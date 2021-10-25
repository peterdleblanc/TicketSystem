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
	
	recSet2 = null;
	cObj = null;
	
	alert(myTotal);
	
}




function LockedTicketCount()
{

	var WorkingTable = new Database_Class('WorkingTable')
	
	WorkingTable.DBName = "DefaultTickets_revised.mdb"
	WorkingTable.QryString = "SELECT * FROM TO_BE_WORKED WHERE BEING_WORK_BY IS NOT NULL AND BEING_WORK_BY <> 'z'";
	
	WorkingTable.LockedRecordCount()
	
	WorkingTable = null;

}

function ViewTicketAgents()
{

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
	UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	
	recSet.close
	cObj.close
	
	recSet = null;
	cObj = null;
	

}

function ViewAgentComments()
{

	var ReportFrom;
	var ReportTo;
	var ReportOn;
		
	ReportOn = document.frmAdmin.txt_EmpNum.value;
	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;

	window.grStr="SELECT TOP 300 * FROM Comment_History WHERE EmpNum='" + ReportOn + "' ORDER BY SortDate DESC"; //AND DATE BETWEEN '" + ReportFrom + "' AND '" + ReportTo +"'";
	
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
	UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	
	recSet.close
	cObj.close
	
	recSet = null;
	cObj = null;


//	var WorkingTable = new Database_Class('WorkingTable')
	
//	WorkingTable.ViewAgentComments()
	
	

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
	UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	
	recSet = null;
	cObj = null;
	


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
	
	recSet = null;
	cObj = null;
	
	

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
	myConnectString += " WHERE Date=#" + TodaysDate +"# AND EmpNum='" + EmpNum + "'";

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
	 
	 
     //document.frmmain.txt_RecLocked.value = LockedCount;
	 document.getElementById('lbl_RecLocked').innerText = LockedCount;
     //document.frmmain.txt_RecCompleted.value = TotalCompleted;
	 document.getElementById('lbl_RecCompleted').innerText = TotalCompleted;

	recSet3.close
	cObj2.close
	
	recSet3 = null;
	cObj2 = null;
	
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
	
	recSet2 = null;
	cObj = null;
	
	
	

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
	myConnectString += " WHERE EmpNum='" + EmpNum + "'" + " AND Date=#" + TodaysDate + "#";
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
		UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	
	
	
     recSet3.close
     cObj2.close

    recSet3 = null;
	cObj2 = null;
	

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
	myConnectString += " WHERE Date BETWEEN #" + ReportFrom + "# AND #" + ReportTo +"#";
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
		UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	 
	recSet3 = null;
	cObj2 = null;
	

    
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
	
	recSet3 = null;
	cObj2 = null;
	
	UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=850,height=450,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
	UserStats.document.write('<html>\n');

	UserStats.document.write('	<head>\n');
	UserStats.document.write('		<!-- CSS STYLESHEET -->\n');
	UserStats.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
	UserStats.document.write('		<!-- JAVASCRIPT LINKS -->\n');
	UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	myConnectString += " WHERE " + "Date BETWEEN #" + ReportFrom + "# AND #" + ReportTo +"#";
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
	
	recSet3 = null;
	cObj2 = null;

}

function DefaultTotalSpreadsheet()
{

    var EmpList=new Array()
	var EmpName=new Array()
	
	
	var UserCount;
	var i;
	var temp;
	
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

	//alert(recSet3.RecordCount)
	
	recSet3.MoveFirst;
	UserCount = 0;
	
	//var GeneratedList = "";
	
	while(recSet3.EOF == 0)
	{
		EmpList[UserCount] = recSet3.Fields('EmpNum').value;
		EmpName[UserCount] = recSet3.Fields('Name').value;
		
		//GeneratedList += EmpList[UserCount]+" \ "+EmpName[UserCount]+"\n";
		
		UserCount++;
		recSet3.MoveNext
	}
	
	//alert(GeneratedList)
	
	recSet3.Close
	cObj2.close
	
	recSet3 = null;
	cObj2 = null;
	
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
		sheet.Cells(myRow,1) = 'Report Name';
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,2) = 'Assign Ticket';
		sheet.Cells(myRow,2).Font.Bold = 'TRUE';
		sheet.Cells(myRow,3) = 'Resolved';
		sheet.Cells(myRow,3).Font.Bold = 'TRUE';
		sheet.Cells(myRow,4) = 'Callbacks Done';
		sheet.Cells(myRow,4).Font.Bold = 'TRUE';
		sheet.Cells(myRow,5) = 'To Self';
		sheet.Cells(myRow,5).Font.Bold = 'TRUE';
		sheet.Cells(myRow,6) = 'To Public';
		sheet.Cells(myRow,6).Font.Bold = 'TRUE';
		sheet.Cells(myRow,7) = 'Reassigned';
		sheet.Cells(myRow,7).Font.Bold = 'TRUE';
		sheet.Cells(myRow,8) = 'Never Worked';
		sheet.Cells(myRow,8).Font.Bold = 'TRUE';
		sheet.Cells(myRow,9) = 'Total Completed';
		sheet.Cells(myRow,9).Font.Bold = 'TRUE';
	
	for(i=0; i<=UserCount; i++)
	{
		ReportOn = EmpList[i];
		ReportName = EmpName[i];
		myRow=CreateAgentSpreadSheet(ReportOn, ReportName, sheet, book, excel, myRow)
	}
	
	temp = myRow
	
	myRow++
	myRow++
		sheet.Cells(myRow,2) = '=SUM(B5:B'+ temp + ')'
		sheet.Cells(myRow,3) = '=SUM(C5:C'+ temp + ')'
		sheet.Cells(myRow,4) = '=SUM(D5:D'+ temp + ')'
		sheet.Cells(myRow,5) = '=SUM(E5:E'+ temp + ')'
		sheet.Cells(myRow,6) = '=SUM(F5:F'+ temp + ')'
		sheet.Cells(myRow,7) = '=SUM(G5:G'+ temp + ')'
		sheet.Cells(myRow,8) = '=SUM(H5:H'+ temp + ')'
		sheet.Cells(myRow,9) = '=SUM(I5:I'+ temp + ')'
	
	excel = null;
	book = null;
	sheet = null;
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
	myConnectString += " WHERE " + "Date BETWEEN #" + document.frmAdmin.txt_From.value + "# AND #" + document.frmAdmin.txt_To.value + "# AND EmpNum='" + EmpNum + "'";
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

		myRow++
		sheet.Cells(myRow,1) = ReportName;
		sheet.Cells(myRow,2) = AssignCount;
		sheet.Cells(myRow,3) = CompletedCount;
		sheet.Cells(myRow,4) = CallbackDone;
		sheet.Cells(myRow,5) = CallbacksSelf;
		sheet.Cells(myRow,6) = CallbacksPublic;
		sheet.Cells(myRow,7) = Reassigned;
		sheet.Cells(myRow,8) = NeverWorked;
		sheet.Cells(myRow,9) = TotalCompleted;
		//myRow++ 
		
	}
	
	recSet3.close
	cObj2.close

	recSet3 = null;
	cObj2 = null;
	
	return myRow;

}

function AgentDaybyDay()
{
	
	var AT10=new Array()
	var AT11=new Array()
	var AT12=new Array()
	var AT13=new Array()
	var AT14=new Array()
	var AT15=new Array()
	var AT16=new Array()
	var AT17=new Array()
	var AT18=new Array()
	var AT19=new Array()
	var AT20=new Array()
	var AT21=new Array()
	var AT22=new Array()
	var AT23=new Array()
	var AT0=new Array()
	var AT1=new Array()
	var AT2=new Array()
	
	//0//var AssignCount;
	//1//var CompletedCount;
	//2//var CallbacksSelf;
	//3//var CallbacksPublic;
	//4//var Reassigned;
	//5//var NeverWorked;
	//6//var CallbackDone;
	//7//var TotalCompleted;

	
	var EmpNum;
	var UserAction;

	var RequestDate;
	var RequestTime;

	var ReportFrom;
	var ReportTo;
	var myConnectString;
	
	var myRow;
	var myCol;

	var excel = new ActiveXObject ("Excel.Application");
	excel.visible = true;
	var book = excel.Workbooks.Add ();
	var sheet = excel.Worksheets(1);

	myRow = 1;
	myCol = 1;
	
	sheet.Columns(1).columnwidth = 14;
	sheet.Columns(2).columnwidth = 14;
	sheet.Columns(3).columnwidth = 14;
	sheet.Columns(4).columnwidth = 14;
	sheet.Columns(5).columnwidth = 14;
	sheet.Columns(6).columnwidth = 14;
	sheet.Columns(7).columnwidth = 14;
	sheet.Columns(8).columnwidth = 14;
	sheet.Columns(9).columnwidth = 14;
	

	for(i=0;i<8;i++)
	{
		AT10[i]=0
		AT11[i]=0
		AT12[i]=0
		AT13[i]=0
		AT14[i]=0
		AT15[i]=0
		AT16[i]=0
		AT17[i]=0
		AT18[i]=0
		AT19[i]=0
		AT20[i]=0
		AT21[i]=0
		AT22[i]=0
		AT23[i]=0
		AT0[i]=0
		AT1[i]=0
		AT2[i]=0
		
	}

	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;
	ReportOn = document.frmAdmin.txt_EmpNum.value;
	
	EmpNum = ReportOn;
		
	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE " + "Date BETWEEN #" + ReportFrom + "# AND #" + ReportTo +"# AND EmpNum='" + EmpNum + "'";
	window.grStr3=myConnectString;

	
	OpenDatabase2()
	OpenTable3()

	if(recSet3.BOF == 0)
	{
		recSet3.MoveFirst
	}
	 
     
     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;
			RequestTime = recSet3.Fields('Time').value;
            switch(UserAction)
            {
            case 'Assign Ticket':
				if(RequestTime == 10)
					AT10[0]++
				if(RequestTime == 11)
					AT11[0]++
				if(RequestTime == 12)
					AT12[0]++
				if(RequestTime == 13)
					AT13[0]++
				if(RequestTime == 14)
					AT14[0]++
				if(RequestTime == 15)
					AT15[0]++
				if(RequestTime == 16)
					AT16[0]++
				if(RequestTime == 17)
					AT17[0]++
				if(RequestTime == 18)
					AT18[0]++
				if(RequestTime == 19)
					AT19[0]++
				if(RequestTime == 20)
					AT20[0]++
				if(RequestTime == 21)
					AT21[0]++
				if(RequestTime == 22)
					AT22[0]++
				if(RequestTime == 23)
					AT23[0]++
				if(RequestTime == 0)
					AT0[0]++
				if(RequestTime == 1)
					AT1[0]++
				if(RequestTime == 2)
					AT2[0]++
                break
				
            case 'Completed':
				if(RequestTime == 10)
					AT10[1]++
				if(RequestTime == 11)
					AT11[1]++
				if(RequestTime == 12)
					AT12[1]++
				if(RequestTime == 13)
					AT13[1]++
				if(RequestTime == 14)
					AT14[1]++
				if(RequestTime == 15)
					AT15[1]++
				if(RequestTime == 16)
					AT16[1]++
				if(RequestTime == 17)
					AT17[1]++
				if(RequestTime == 18)
					AT18[1]++
				if(RequestTime == 19)
					AT19[1]++
				if(RequestTime == 20)
					AT20[1]++
				if(RequestTime == 21)
					AT21[1]++
				if(RequestTime == 22)
					AT22[1]++
				if(RequestTime == 23)
					AT23[1]++
				if(RequestTime == 0)
					AT0[1]++
				if(RequestTime == 1)
					AT1[1]++
				if(RequestTime == 2)
					AT2[1]++
                break
				
			case 'ToSelf':
				if(RequestTime == 10)
					AT10[2]++
				if(RequestTime == 11)
					AT11[2]++
				if(RequestTime == 12)
					AT12[2]++
				if(RequestTime == 13)
					AT13[2]++
				if(RequestTime == 14)
					AT14[2]++
				if(RequestTime == 15)
					AT15[2]++
				if(RequestTime == 16)
					AT16[2]++
				if(RequestTime == 17)
					AT17[2]++
				if(RequestTime == 18)
					AT18[2]++
				if(RequestTime == 19)
					AT19[2]++
				if(RequestTime == 20)
					AT20[2]++
				if(RequestTime == 21)
					AT21[2]++
				if(RequestTime == 22)
					AT22[2]++
				if(RequestTime == 23)
					AT23[2]++
				if(RequestTime == 0)
					AT0[2]++
				if(RequestTime == 1)
					AT1[2]++
				if(RequestTime == 2)
					AT2[2]++
				break
				
			case 'ToPublic':
				if(RequestTime == 10)
					AT10[3]++
				if(RequestTime == 11)
					AT11[3]++
				if(RequestTime == 12)
					AT12[3]++
				if(RequestTime == 13)
					AT13[3]++
				if(RequestTime == 14)
					AT14[3]++
				if(RequestTime == 15)
					AT15[3]++
				if(RequestTime == 16)
					AT16[3]++
				if(RequestTime == 17)
					AT17[3]++
				if(RequestTime == 18)
					AT18[3]++
				if(RequestTime == 19)
					AT19[3]++
				if(RequestTime == 20)
					AT20[3]++
				if(RequestTime == 21)
					AT21[3]++
				if(RequestTime == 22)
					AT22[3]++
				if(RequestTime == 23)
					AT23[3]++
				if(RequestTime == 0)
					AT0[3]++
				if(RequestTime == 1)
					AT1[3]++
				if(RequestTime == 2)
					AT2[3]++
				break
				
			case 'Reassigned':
				if(RequestTime == 10)
					AT10[4]++
				if(RequestTime == 11)
					AT11[4]++
				if(RequestTime == 12)
					AT12[4]++
				if(RequestTime == 13)
					AT13[4]++
				if(RequestTime == 14)
					AT14[4]++
				if(RequestTime == 15)
					AT15[4]++
				if(RequestTime == 16)
					AT16[4]++
				if(RequestTime == 17)
					AT17[4]++
				if(RequestTime == 18)
					AT18[4]++
				if(RequestTime == 19)
					AT19[4]++
				if(RequestTime == 20)
					AT20[4]++
				if(RequestTime == 21)
					AT21[4]++
				if(RequestTime == 22)
					AT22[4]++
				if(RequestTime == 23)
					AT23[4]++
				if(RequestTime == 0)
					AT0[4]++
				if(RequestTime == 1)
					AT1[4]++
				if(RequestTime == 2)
					AT2[4]++
				break
				
			case 'NeverWorked':
				if(RequestTime == 10)
					AT10[5]++
				if(RequestTime == 11)
					AT11[5]++
				if(RequestTime == 12)
					AT12[5]++
				if(RequestTime == 13)
					AT13[5]++
				if(RequestTime == 14)
					AT14[5]++
				if(RequestTime == 15)
					AT15[5]++
				if(RequestTime == 16)
					AT16[5]++
				if(RequestTime == 17)
					AT17[5]++
				if(RequestTime == 18)
					AT18[5]++
				if(RequestTime == 19)
					AT19[5]++
				if(RequestTime == 20)
					AT20[5]++
				if(RequestTime == 21)
					AT21[5]++
				if(RequestTime == 22)
					AT22[5]++
				if(RequestTime == 23)
					AT23[5]++
				if(RequestTime == 0)
					AT0[5]++
				if(RequestTime == 1)
					AT1[5]++
				if(RequestTime == 2)
					AT2[5]++
				break
			case 'DidCallBack':
				if(RequestTime == 10)
					AT10[6]++
				if(RequestTime == 11)
					AT11[6]++
				if(RequestTime == 12)
					AT12[6]++
				if(RequestTime == 13)
					AT13[6]++
				if(RequestTime == 14)
					AT14[6]++
				if(RequestTime == 15)
					AT15[6]++
				if(RequestTime == 16)
					AT16[6]++
				if(RequestTime == 17)
					AT17[6]++
				if(RequestTime == 18)
					AT18[6]++
				if(RequestTime == 19)
					AT19[6]++
				if(RequestTime == 20)
					AT20[6]++
				if(RequestTime == 21)
					AT21[6]++
				if(RequestTime == 22)
					AT22[6]++
				if(RequestTime == 23)
					AT23[6]++
				if(RequestTime == 0)
					AT0[6]++
				if(RequestTime == 1)
					AT1[6]++
				if(RequestTime == 2)
					AT2[6]++
				break
            default:
            }
        }
        recSet3.MoveNext;
     }
	 
	//0//var AssignCount;
	//1//var CompletedCount;
	//2//var CallbacksSelf;
	//3//var CallbacksPublic;
	//4//var Reassigned;
	//5//var NeverWorked;
	//6//var CallbackDone;
	//7//var TotalCompleted;

		sheet.Cells(myRow,1) = EmpNum;
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,1).Font.Size = 12;
		sheet.Cells(myRow,1).Font.ColorIndex = 5;
		myRow++;
		
		sheet.Cells(myRow,1) = 'Time';
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,2) = 'Assign Ticket';
		sheet.Cells(myRow,2).Font.Bold = 'TRUE';
		sheet.Cells(myRow,3) = 'Resolved';
		sheet.Cells(myRow,3).Font.Bold = 'TRUE';
		sheet.Cells(myRow,4) = 'Callbacks Self';
		sheet.Cells(myRow,4).Font.Bold = 'TRUE';
		sheet.Cells(myRow,5) = 'Callbacks Public';
		sheet.Cells(myRow,5).Font.Bold = 'TRUE';
		sheet.Cells(myRow,6) = 'Reassigned';
		sheet.Cells(myRow,6).Font.Bold = 'TRUE';
		sheet.Cells(myRow,7) = 'Never Worked';
		sheet.Cells(myRow,7).Font.Bold = 'TRUE';
		sheet.Cells(myRow,8) = 'Callback Done';
		sheet.Cells(myRow,8).Font.Bold = 'TRUE';
		sheet.Cells(myRow,9) = 'Total Worked';
		sheet.Cells(myRow,9).Font.Bold = 'TRUE';
	
		myRow++
		sheet.Cells(myRow,1) = '11 AM';
		sheet.Cells(myRow,2) = AT11[0]
		sheet.Cells(myRow,3) = AT11[1]
		sheet.Cells(myRow,4) = AT11[2]
		sheet.Cells(myRow,5) = AT11[3]
		sheet.Cells(myRow,6) = AT11[4]
		sheet.Cells(myRow,7) = AT11[5]
		sheet.Cells(myRow,8) = AT11[6]
		sheet.Cells(myRow,9) = AT11[7]
		myRow++
		sheet.Cells(myRow,1) = '12 PM';
		sheet.Cells(myRow,2) = AT12[0]
		sheet.Cells(myRow,3) = AT12[1]
		sheet.Cells(myRow,4) = AT12[2]
		sheet.Cells(myRow,5) = AT12[3]
		sheet.Cells(myRow,6) = AT12[4]
		sheet.Cells(myRow,7) = AT12[5]
		sheet.Cells(myRow,8) = AT12[6]
		sheet.Cells(myRow,9) = AT12[7]
		myRow++
		sheet.Cells(myRow,1) = '1 PM';
		sheet.Cells(myRow,2) = AT13[0]
		sheet.Cells(myRow,3) = AT13[1]
		sheet.Cells(myRow,4) = AT13[2]
		sheet.Cells(myRow,5) = AT13[3]
		sheet.Cells(myRow,6) = AT13[4]
		sheet.Cells(myRow,7) = AT13[5]
		sheet.Cells(myRow,8) = AT13[6]
		sheet.Cells(myRow,9) = AT13[7]
		myRow++
		sheet.Cells(myRow,1) = '2 PM';
		sheet.Cells(myRow,2) = AT14[0]
		sheet.Cells(myRow,3) = AT14[1]
		sheet.Cells(myRow,4) = AT14[2]
		sheet.Cells(myRow,5) = AT14[3]
		sheet.Cells(myRow,6) = AT14[4]
		sheet.Cells(myRow,7) = AT14[5]
		sheet.Cells(myRow,8) = AT14[6]
		sheet.Cells(myRow,9) = AT14[7]
		myRow++
		sheet.Cells(myRow,1) = '3 PM';
		sheet.Cells(myRow,2) = AT15[0]
		sheet.Cells(myRow,3) = AT15[1]
		sheet.Cells(myRow,4) = AT15[2]
		sheet.Cells(myRow,5) = AT15[3]
		sheet.Cells(myRow,6) = AT15[4]
		sheet.Cells(myRow,7) = AT15[5]
		sheet.Cells(myRow,8) = AT15[6]
		sheet.Cells(myRow,9) = AT15[7]
		myRow++
		sheet.Cells(myRow,1) = '4 PM';
		sheet.Cells(myRow,2) = AT16[0]
		sheet.Cells(myRow,3) = AT16[1]
		sheet.Cells(myRow,4) = AT16[2]
		sheet.Cells(myRow,5) = AT16[3]
		sheet.Cells(myRow,6) = AT16[4]
		sheet.Cells(myRow,7) = AT16[5]
		sheet.Cells(myRow,8) = AT16[6]
		sheet.Cells(myRow,9) = AT16[7]
		myRow++
		sheet.Cells(myRow,1) = '5 PM';
		sheet.Cells(myRow,2) = AT17[0]
		sheet.Cells(myRow,3) = AT17[1]
		sheet.Cells(myRow,4) = AT17[2]
		sheet.Cells(myRow,5) = AT17[3]
		sheet.Cells(myRow,6) = AT17[4]
		sheet.Cells(myRow,7) = AT17[5]
		sheet.Cells(myRow,8) = AT17[6]
		sheet.Cells(myRow,9) = AT17[7]
		myRow++
		sheet.Cells(myRow,1) = '6 PM';
		sheet.Cells(myRow,2) = AT18[0]
		sheet.Cells(myRow,3) = AT18[1]
		sheet.Cells(myRow,4) = AT18[2]
		sheet.Cells(myRow,5) = AT18[3]
		sheet.Cells(myRow,6) = AT18[4]
		sheet.Cells(myRow,7) = AT18[5]
		sheet.Cells(myRow,8) = AT18[6]
		sheet.Cells(myRow,9) = AT18[7]
		myRow++
		sheet.Cells(myRow,1) = '7 PM';
		sheet.Cells(myRow,2) = AT19[0]
		sheet.Cells(myRow,3) = AT19[1]
		sheet.Cells(myRow,4) = AT19[2]
		sheet.Cells(myRow,5) = AT19[3]
		sheet.Cells(myRow,6) = AT19[4]
		sheet.Cells(myRow,7) = AT19[5]
		sheet.Cells(myRow,8) = AT19[6]
		sheet.Cells(myRow,9) = AT19[7]
		myRow++
		sheet.Cells(myRow,1) = '8 PM';
		sheet.Cells(myRow,2) = AT20[0]
		sheet.Cells(myRow,3) = AT20[1]
		sheet.Cells(myRow,4) = AT20[2]
		sheet.Cells(myRow,5) = AT20[3]
		sheet.Cells(myRow,6) = AT20[4]
		sheet.Cells(myRow,7) = AT20[5]
		sheet.Cells(myRow,8) = AT20[6]
		sheet.Cells(myRow,9) = AT20[7]
		myRow++
		sheet.Cells(myRow,1) = '9 PM';
		sheet.Cells(myRow,2) = AT21[0]
		sheet.Cells(myRow,3) = AT21[1]
		sheet.Cells(myRow,4) = AT21[2]
		sheet.Cells(myRow,5) = AT21[3]
		sheet.Cells(myRow,6) = AT21[4]
		sheet.Cells(myRow,7) = AT21[5]
		sheet.Cells(myRow,8) = AT21[6]
		sheet.Cells(myRow,9) = AT21[7]
		myRow++
		sheet.Cells(myRow,1) = '10 PM';
		sheet.Cells(myRow,2) = AT22[0]
		sheet.Cells(myRow,3) = AT22[1]
		sheet.Cells(myRow,4) = AT22[2]
		sheet.Cells(myRow,5) = AT22[3]
		sheet.Cells(myRow,6) = AT22[4]
		sheet.Cells(myRow,7) = AT22[5]
		sheet.Cells(myRow,8) = AT22[6]
		sheet.Cells(myRow,9) = AT22[7]
		myRow++
		sheet.Cells(myRow,1) = '11 PM';
		sheet.Cells(myRow,2) = AT23[0]
		sheet.Cells(myRow,3) = AT23[1]
		sheet.Cells(myRow,4) = AT23[2]
		sheet.Cells(myRow,5) = AT23[3]
		sheet.Cells(myRow,6) = AT23[4]
		sheet.Cells(myRow,7) = AT23[5]
		sheet.Cells(myRow,8) = AT23[6]
		sheet.Cells(myRow,9) = AT23[7]
		myRow++
		sheet.Cells(myRow,1) = '12 AM';
		sheet.Cells(myRow,2) = AT0[0]
		sheet.Cells(myRow,3) = AT0[1]
		sheet.Cells(myRow,4) = AT0[2]
		sheet.Cells(myRow,5) = AT0[3]
		sheet.Cells(myRow,6) = AT0[4]
		sheet.Cells(myRow,7) = AT0[5]
		sheet.Cells(myRow,8) = AT0[6]
		sheet.Cells(myRow,9) = AT0[7]
		myRow++
		sheet.Cells(myRow,1) = '1 AM';
		sheet.Cells(myRow,2) = AT1[0]
		sheet.Cells(myRow,3) = AT1[1]
		sheet.Cells(myRow,4) = AT1[2]
		sheet.Cells(myRow,5) = AT1[3]
		sheet.Cells(myRow,6) = AT1[4]
		sheet.Cells(myRow,7) = AT1[5]
		sheet.Cells(myRow,8) = AT1[6]
		sheet.Cells(myRow,9) = AT1[7]
	
	recSet3.close
	cObj2.close

	recSet3 = null;
	cObj2 = null;
	
	return myRow;
	

}

function TeamDayByDaySpreadsheet()
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
	cObj2.close
	
	recSet3 = null;
	cObj2 = null;
	
	alert("Creating Spreadsheet")

	var myRow;
	var myCol;

	var excel = new ActiveXObject ("Excel.Application");
	excel.visible = true;
	var book = excel.Workbooks.Add ();
	var sheet = excel.Worksheets(1);

	myRow = 1;
	myCol = 1;
	
	sheet.Columns(1).columnwidth = 14;
	sheet.Columns(2).columnwidth = 14;
	sheet.Columns(3).columnwidth = 14;
	sheet.Columns(4).columnwidth = 14;
	sheet.Columns(5).columnwidth = 14;
	sheet.Columns(6).columnwidth = 14;
	sheet.Columns(7).columnwidth = 14;
	sheet.Columns(8).columnwidth = 14;
	sheet.Columns(9).columnwidth = 14;

	
	for(i=0; i<=UserCount; i++)
	{
		ReportOn = EmpList[i];
		ReportName = EmpName[i];
		myRow=TeamDaybyDay(ReportOn, sheet, book, excel, myRow)
	}
		

}


function TeamDaybyDay(ReportOn, sheet, book, excel, myRow)
{
	var AT11=new Array()
	var AT12=new Array()
	var AT13=new Array()
	var AT14=new Array()
	var AT15=new Array()
	var AT16=new Array()
	var AT17=new Array()
	var AT18=new Array()
	var AT19=new Array()
	var AT20=new Array()
	var AT21=new Array()
	var AT22=new Array()
	var AT23=new Array()
	var AT0=new Array()
	var AT1=new Array()
	var AT2=new Array()
	
	//0//var AssignCount;
	//1//var CompletedCount;
	//2//var CallbacksSelf;
	//3//var CallbacksPublic;
	//4//var Reassigned;
	//5//var NeverWorked;
	//6//var CallbackDone;
	//7//var TotalCompleted;

	var WorkDone;
	var EmpNum;
	var UserAction;

	var RequestDate;
	var RequestTime;

	var ReportFrom;
	var ReportTo;
	var myConnectString;
	
	var temp1;
	var temp2;


	EmpNum = ReportOn
	
	for(i=0;i<8;i++)
	{
		AT11[i]=0
		AT12[i]=0
		AT13[i]=0
		AT14[i]=0
		AT15[i]=0
		AT16[i]=0
		AT17[i]=0
		AT18[i]=0
		AT19[i]=0
		AT20[i]=0
		AT21[i]=0
		AT22[i]=0
		AT23[i]=0
		AT0[i]=0
		AT1[i]=0
		AT2[i]=0
		
	}

	ReportFrom = document.frmAdmin.txt_From.value;
	ReportTo = document.frmAdmin.txt_To.value;

	
	EmpNum = ReportOn;
		
	myConnectString = "SELECT * FROM DefaultTicketStats";
	myConnectString += " WHERE " + "Date BETWEEN #" + ReportFrom + "# AND #" + ReportTo +"# AND EmpNum='" + EmpNum + "'";
	window.grStr3=myConnectString;
	
	OpenDatabase2()
	OpenTable3()

	if(recSet3.BOF == 0)
	{
		recSet3.MoveFirst
	}
	 
     
     while(recSet3.EOF == 0)
     {
        if(EmpNum == recSet3.Fields('EmpNum').value)
        {
            UserAction = recSet3.Fields('Action').value;
			RequestTime = recSet3.Fields('Time').value;
            switch(UserAction)
            {
            case 'Assign Ticket':
				if(RequestTime == 11)
					AT11[0]++
				if(RequestTime == 12)
					AT12[0]++
				if(RequestTime == 13)
					AT13[0]++
				if(RequestTime == 14)
					AT14[0]++
				if(RequestTime == 15)
					AT15[0]++
				if(RequestTime == 16)
					AT16[0]++
				if(RequestTime == 17)
					AT17[0]++
				if(RequestTime == 18)
					AT18[0]++
				if(RequestTime == 19)
					AT19[0]++
				if(RequestTime == 20)
					AT20[0]++
				if(RequestTime == 21)
					AT21[0]++
				if(RequestTime == 22)
					AT22[0]++
				if(RequestTime == 23)
					AT23[0]++
				if(RequestTime == 0)
					AT0[0]++
				if(RequestTime == 1)
					AT1[0]++
				if(RequestTime == 2)
					AT2[0]++
                break
				
            case 'Completed':
				if(RequestTime == 11)
					AT11[1]++
				if(RequestTime == 12)
					AT12[1]++
				if(RequestTime == 13)
					AT13[1]++
				if(RequestTime == 14)
					AT14[1]++
				if(RequestTime == 15)
					AT15[1]++
				if(RequestTime == 16)
					AT16[1]++
				if(RequestTime == 17)
					AT17[1]++
				if(RequestTime == 18)
					AT18[1]++
				if(RequestTime == 19)
					AT19[1]++
				if(RequestTime == 20)
					AT20[1]++
				if(RequestTime == 21)
					AT21[1]++
				if(RequestTime == 22)
					AT22[1]++
				if(RequestTime == 23)
					AT23[1]++
				if(RequestTime == 0)
					AT0[1]++
				if(RequestTime == 1)
					AT1[1]++
				if(RequestTime == 2)
					AT2[1]++
                break
				
			case 'ToSelf':
				if(RequestTime == 11)
					AT11[2]++
				if(RequestTime == 12)
					AT12[2]++
				if(RequestTime == 13)
					AT13[2]++
				if(RequestTime == 14)
					AT14[2]++
				if(RequestTime == 15)
					AT15[2]++
				if(RequestTime == 16)
					AT16[2]++
				if(RequestTime == 17)
					AT17[2]++
				if(RequestTime == 18)
					AT18[2]++
				if(RequestTime == 19)
					AT19[2]++
				if(RequestTime == 20)
					AT20[2]++
				if(RequestTime == 21)
					AT21[2]++
				if(RequestTime == 22)
					AT22[2]++
				if(RequestTime == 23)
					AT23[2]++
				if(RequestTime == 0)
					AT0[2]++
				if(RequestTime == 1)
					AT1[2]++
				if(RequestTime == 2)
					AT2[2]++
				break
				
			case 'ToPublic':
				if(RequestTime == 11)
					AT11[3]++
				if(RequestTime == 12)
					AT12[3]++
				if(RequestTime == 13)
					AT13[3]++
				if(RequestTime == 14)
					AT14[3]++
				if(RequestTime == 15)
					AT15[3]++
				if(RequestTime == 16)
					AT16[3]++
				if(RequestTime == 17)
					AT17[3]++
				if(RequestTime == 18)
					AT18[3]++
				if(RequestTime == 19)
					AT19[3]++
				if(RequestTime == 20)
					AT20[3]++
				if(RequestTime == 21)
					AT21[3]++
				if(RequestTime == 22)
					AT22[3]++
				if(RequestTime == 23)
					AT23[3]++
				if(RequestTime == 0)
					AT0[3]++
				if(RequestTime == 1)
					AT1[3]++
				if(RequestTime == 2)
					AT2[3]++
				break
				
			case 'Reassigned':
				if(RequestTime == 11)
					AT11[4]++
				if(RequestTime == 12)
					AT12[4]++
				if(RequestTime == 13)
					AT13[4]++
				if(RequestTime == 14)
					AT14[4]++
				if(RequestTime == 15)
					AT15[4]++
				if(RequestTime == 16)
					AT16[4]++
				if(RequestTime == 17)
					AT17[4]++
				if(RequestTime == 18)
					AT18[4]++
				if(RequestTime == 19)
					AT19[4]++
				if(RequestTime == 20)
					AT20[4]++
				if(RequestTime == 21)
					AT21[4]++
				if(RequestTime == 22)
					AT22[4]++
				if(RequestTime == 23)
					AT23[4]++
				if(RequestTime == 0)
					AT0[4]++
				if(RequestTime == 1)
					AT1[4]++
				if(RequestTime == 2)
					AT2[4]++
				break
				
			case 'NeverWorked':
				if(RequestTime == 11)
					AT11[5]++
				if(RequestTime == 12)
					AT12[5]++
				if(RequestTime == 13)
					AT13[5]++
				if(RequestTime == 14)
					AT14[5]++
				if(RequestTime == 15)
					AT15[5]++
				if(RequestTime == 16)
					AT16[5]++
				if(RequestTime == 17)
					AT17[5]++
				if(RequestTime == 18)
					AT18[5]++
				if(RequestTime == 19)
					AT19[5]++
				if(RequestTime == 20)
					AT20[5]++
				if(RequestTime == 21)
					AT21[5]++
				if(RequestTime == 22)
					AT22[5]++
				if(RequestTime == 23)
					AT23[5]++
				if(RequestTime == 0)
					AT0[5]++
				if(RequestTime == 1)
					AT1[5]++
				if(RequestTime == 2)
					AT2[5]++
				break
			case 'DidCallBack':
				if(RequestTime == 11)
					AT11[6]++
				if(RequestTime == 12)
					AT12[6]++
				if(RequestTime == 13)
					AT13[6]++
				if(RequestTime == 14)
					AT14[6]++
				if(RequestTime == 15)
					AT15[6]++
				if(RequestTime == 16)
					AT16[6]++
				if(RequestTime == 17)
					AT17[6]++
				if(RequestTime == 18)
					AT18[6]++
				if(RequestTime == 19)
					AT19[6]++
				if(RequestTime == 20)
					AT20[6]++
				if(RequestTime == 21)
					AT21[6]++
				if(RequestTime == 22)
					AT22[6]++
				if(RequestTime == 23)
					AT23[6]++
				if(RequestTime == 0)
					AT0[6]++
				if(RequestTime == 1)
					AT1[6]++
				if(RequestTime == 2)
					AT2[6]++
				break
            default:
            }
        }
        recSet3.MoveNext;
     }
	 
	 WorkDone=AT11[0] + AT12[0] + AT13[0] + AT14[0] + AT15[0] + AT16[0] + AT17[0] + AT18[0] + AT19[0] + AT20[0] + AT21[0] + AT22[0] + AT23[0] + AT0[0] + AT1[0] + AT2[0]
	 
	//0//var AssignCount;
	//1//var CompletedCount;
	//2//var CallbacksSelf;
	//3//var CallbacksPublic;
	//4//var Reassigned;
	//5//var NeverWorked;
	//6//var CallbackDone;
	//7//var TotalCompleted;
	
	if(WorkDone != 0)
	{
		myRow++;

		sheet.Cells(myRow,1) = EmpNum;
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,1).Font.Size = 12;
		sheet.Cells(myRow,1).Font.ColorIndex = 5;
		myRow++;
		
		sheet.Cells(myRow,1) = 'Time';
		sheet.Cells(myRow,1).Font.Bold = 'TRUE';
		sheet.Cells(myRow,2) = 'Assign Ticket';
		sheet.Cells(myRow,2).Font.Bold = 'TRUE';
		sheet.Cells(myRow,3) = 'Resolved';
		sheet.Cells(myRow,3).Font.Bold = 'TRUE';
		sheet.Cells(myRow,4) = 'Callbacks Self';
		sheet.Cells(myRow,4).Font.Bold = 'TRUE';
		sheet.Cells(myRow,5) = 'Callbacks Public';
		sheet.Cells(myRow,5).Font.Bold = 'TRUE';
		sheet.Cells(myRow,6) = 'Reassigned';
		sheet.Cells(myRow,6).Font.Bold = 'TRUE';
		sheet.Cells(myRow,7) = 'Never Worked';
		sheet.Cells(myRow,7).Font.Bold = 'TRUE';
		sheet.Cells(myRow,8) = 'Callback Done';
		sheet.Cells(myRow,8).Font.Bold = 'TRUE';
		sheet.Cells(myRow,9) = 'Total Worked';
		sheet.Cells(myRow,9).Font.Bold = 'TRUE';
	
		myRow++
		temp1 = myRow;
		sheet.Cells(myRow,1) = '11 AM';
		sheet.Cells(myRow,2) = AT11[0]
		sheet.Cells(myRow,3) = AT11[1]
		sheet.Cells(myRow,4) = AT11[2]
		sheet.Cells(myRow,5) = AT11[3]
		sheet.Cells(myRow,6) = AT11[4]
		sheet.Cells(myRow,7) = AT11[5]
		sheet.Cells(myRow,8) = AT11[6]
		sheet.Cells(myRow,9) = AT11[7]
		myRow++
		sheet.Cells(myRow,1) = '12 PM';
		sheet.Cells(myRow,2) = AT12[0]
		sheet.Cells(myRow,3) = AT12[1]
		sheet.Cells(myRow,4) = AT12[2]
		sheet.Cells(myRow,5) = AT12[3]
		sheet.Cells(myRow,6) = AT12[4]
		sheet.Cells(myRow,7) = AT12[5]
		sheet.Cells(myRow,8) = AT12[6]
		sheet.Cells(myRow,9) = AT12[7]
		myRow++
		sheet.Cells(myRow,1) = '1 PM';
		sheet.Cells(myRow,2) = AT13[0]
		sheet.Cells(myRow,3) = AT13[1]
		sheet.Cells(myRow,4) = AT13[2]
		sheet.Cells(myRow,5) = AT13[3]
		sheet.Cells(myRow,6) = AT13[4]
		sheet.Cells(myRow,7) = AT13[5]
		sheet.Cells(myRow,8) = AT13[6]
		sheet.Cells(myRow,9) = AT13[7]
		myRow++
		sheet.Cells(myRow,1) = '2 PM';
		sheet.Cells(myRow,2) = AT14[0]
		sheet.Cells(myRow,3) = AT14[1]
		sheet.Cells(myRow,4) = AT14[2]
		sheet.Cells(myRow,5) = AT14[3]
		sheet.Cells(myRow,6) = AT14[4]
		sheet.Cells(myRow,7) = AT14[5]
		sheet.Cells(myRow,8) = AT14[6]
		sheet.Cells(myRow,9) = AT14[7]
		myRow++
		sheet.Cells(myRow,1) = '3 PM';
		sheet.Cells(myRow,2) = AT15[0]
		sheet.Cells(myRow,3) = AT15[1]
		sheet.Cells(myRow,4) = AT15[2]
		sheet.Cells(myRow,5) = AT15[3]
		sheet.Cells(myRow,6) = AT15[4]
		sheet.Cells(myRow,7) = AT15[5]
		sheet.Cells(myRow,8) = AT15[6]
		sheet.Cells(myRow,9) = AT15[7]
		myRow++
		sheet.Cells(myRow,1) = '4 PM';
		sheet.Cells(myRow,2) = AT16[0]
		sheet.Cells(myRow,3) = AT16[1]
		sheet.Cells(myRow,4) = AT16[2]
		sheet.Cells(myRow,5) = AT16[3]
		sheet.Cells(myRow,6) = AT16[4]
		sheet.Cells(myRow,7) = AT16[5]
		sheet.Cells(myRow,8) = AT16[6]
		sheet.Cells(myRow,9) = AT16[7]
		myRow++
		sheet.Cells(myRow,1) = '5 PM';
		sheet.Cells(myRow,2) = AT17[0]
		sheet.Cells(myRow,3) = AT17[1]
		sheet.Cells(myRow,4) = AT17[2]
		sheet.Cells(myRow,5) = AT17[3]
		sheet.Cells(myRow,6) = AT17[4]
		sheet.Cells(myRow,7) = AT17[5]
		sheet.Cells(myRow,8) = AT17[6]
		sheet.Cells(myRow,9) = AT17[7]
		myRow++
		sheet.Cells(myRow,1) = '6 PM';
		sheet.Cells(myRow,2) = AT18[0]
		sheet.Cells(myRow,3) = AT18[1]
		sheet.Cells(myRow,4) = AT18[2]
		sheet.Cells(myRow,5) = AT18[3]
		sheet.Cells(myRow,6) = AT18[4]
		sheet.Cells(myRow,7) = AT18[5]
		sheet.Cells(myRow,8) = AT18[6]
		sheet.Cells(myRow,9) = AT18[7]
		myRow++
		sheet.Cells(myRow,1) = '7 PM';
		sheet.Cells(myRow,2) = AT19[0]
		sheet.Cells(myRow,3) = AT19[1]
		sheet.Cells(myRow,4) = AT19[2]
		sheet.Cells(myRow,5) = AT19[3]
		sheet.Cells(myRow,6) = AT19[4]
		sheet.Cells(myRow,7) = AT19[5]
		sheet.Cells(myRow,8) = AT19[6]
		sheet.Cells(myRow,9) = AT19[7]
		myRow++
		sheet.Cells(myRow,1) = '8 PM';
		sheet.Cells(myRow,2) = AT20[0]
		sheet.Cells(myRow,3) = AT20[1]
		sheet.Cells(myRow,4) = AT20[2]
		sheet.Cells(myRow,5) = AT20[3]
		sheet.Cells(myRow,6) = AT20[4]
		sheet.Cells(myRow,7) = AT20[5]
		sheet.Cells(myRow,8) = AT20[6]
		sheet.Cells(myRow,9) = AT20[7]
		myRow++
		sheet.Cells(myRow,1) = '9 PM';
		sheet.Cells(myRow,2) = AT21[0]
		sheet.Cells(myRow,3) = AT21[1]
		sheet.Cells(myRow,4) = AT21[2]
		sheet.Cells(myRow,5) = AT21[3]
		sheet.Cells(myRow,6) = AT21[4]
		sheet.Cells(myRow,7) = AT21[5]
		sheet.Cells(myRow,8) = AT21[6]
		sheet.Cells(myRow,9) = AT21[7]
		myRow++
		sheet.Cells(myRow,1) = '10 PM';
		sheet.Cells(myRow,2) = AT22[0]
		sheet.Cells(myRow,3) = AT22[1]
		sheet.Cells(myRow,4) = AT22[2]
		sheet.Cells(myRow,5) = AT22[3]
		sheet.Cells(myRow,6) = AT22[4]
		sheet.Cells(myRow,7) = AT22[5]
		sheet.Cells(myRow,8) = AT22[6]
		sheet.Cells(myRow,9) = AT22[7]
		myRow++
		sheet.Cells(myRow,1) = '11 PM';
		sheet.Cells(myRow,2) = AT23[0]
		sheet.Cells(myRow,3) = AT23[1]
		sheet.Cells(myRow,4) = AT23[2]
		sheet.Cells(myRow,5) = AT23[3]
		sheet.Cells(myRow,6) = AT23[4]
		sheet.Cells(myRow,7) = AT23[5]
		sheet.Cells(myRow,8) = AT23[6]
		sheet.Cells(myRow,9) = AT23[7]
		myRow++
		sheet.Cells(myRow,1) = '12 AM';
		sheet.Cells(myRow,2) = AT0[0]
		sheet.Cells(myRow,3) = AT0[1]
		sheet.Cells(myRow,4) = AT0[2]
		sheet.Cells(myRow,5) = AT0[3]
		sheet.Cells(myRow,6) = AT0[4]
		sheet.Cells(myRow,7) = AT0[5]
		sheet.Cells(myRow,8) = AT0[6]
		sheet.Cells(myRow,9) = AT0[7]
		myRow++
		sheet.Cells(myRow,1) = '1 AM';
		sheet.Cells(myRow,2) = AT1[0]
		sheet.Cells(myRow,3) = AT1[1]
		sheet.Cells(myRow,4) = AT1[2]
		sheet.Cells(myRow,5) = AT1[3]
		sheet.Cells(myRow,6) = AT1[4]
		sheet.Cells(myRow,7) = AT1[5]
		sheet.Cells(myRow,8) = AT1[6]
		sheet.Cells(myRow,9) = AT1[7]
		
		temp2 = myRow;
		myRow++
		sheet.Cells(myRow,1).Font.ColorIndex = 1;
		sheet.Cells(myRow,1) = 'Totals';
		sheet.Cells(myRow,2).Font.ColorIndex = 3;
		sheet.Cells(myRow,2) = '=SUM(B'+ temp1 + ':B'+ temp2 + ')'
		sheet.Cells(myRow,3).Font.ColorIndex = 3;
		sheet.Cells(myRow,3) = '=SUM(C'+ temp1 + ':C'+ temp2 + ')'
		sheet.Cells(myRow,4).Font.ColorIndex = 3;
		sheet.Cells(myRow,4) = '=SUM(D'+ temp1 + ':D'+ temp2 + ')'
		sheet.Cells(myRow,5).Font.ColorIndex = 3;
		sheet.Cells(myRow,5) = '=SUM(E'+ temp1 + ':E'+ temp2 + ')'
		sheet.Cells(myRow,6).Font.ColorIndex = 3;
		sheet.Cells(myRow,6) = '=SUM(F'+ temp1 + ':F'+ temp2 + ')'
		sheet.Cells(myRow,7).Font.ColorIndex = 3;
		sheet.Cells(myRow,7) = '=SUM(G'+ temp1 + ':G'+ temp2 + ')'
		sheet.Cells(myRow,8).Font.ColorIndex = 3;
		sheet.Cells(myRow,8) = '=SUM(H'+ temp1 + ':H'+ temp2 + ')'
		sheet.Cells(myRow,9).Font.ColorIndex = 3;
		sheet.Cells(myRow,9) = '=SUM(I'+ temp1 + ':I'+ temp2 + ')'
	}
	recSet3.close
	cObj2.close

	recSet3 = null;
	cObj2 = null;
	
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
	
	var SQL_Query = "SELECT * FROM Completed WHERE INVALID_TICKET=true AND WORKED_DATE BETWEEN #" + ReportFrom + "# AND #" + ReportTo +"# ORDER BY WORKED_DATE";
	window.grStr= SQL_Query;
	
	OpenDatabase()
	OpenTable()
	
	sheet.Cells(myRow,1) = 'Proformed Query:';
	sheet.Cells(myRow,1).Font.Bold = 'TRUE';
	sheet.Cells(myRow,1).Font.Size = 12;
	myRow++;
	
	sheet.Cells(myRow,1) = SQL_Query;
	sheet.Cells(myRow,1).Font.Bold = 'FALSE';
	sheet.Cells(myRow,1).Font.Size = 12;
	myRow++;
	
	sheet.Cells(myRow,1) = 'Found: '+recSet.RecordCount;
	sheet.Cells(myRow,1).Font.Bold = 'TRUE';
	sheet.Cells(myRow,1).Font.Size = 12;
	myRow++;
	
	sheet.Cells(myRow,1) = 'Invalid Tickets';
	sheet.Cells(myRow,1).Font.Bold = 'TRUE';
	sheet.Cells(myRow,1).Font.Size = 15;
	myRow++;
	myRow++;
	
	sheet.Columns(1).columnwidth = 12.86;
	sheet.Columns(2).columnwidth = 13;
	sheet.Columns(3).columnwidth = 30.14;
	sheet.Columns(4).columnwidth = 58.29;
	sheet.Columns(5).columnwidth = 9.43;
	
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
	sheet.Cells(myRow,5) = 'Worked';
	sheet.Cells(myRow,5).Font.Bold = 'TRUE';
	//sheet.Cells(myRow,6) = 'Invalid';
	//sheet.Cells(myRow,6).Font.Bold = 'TRUE';
	
	
	myRow++;
	while(recSet.EOF == 0)
	{
		sheet.Cells(myRow,1) = recSet.Fields('TICKET_ID').value;
		sheet.Cells(myRow,2) = recSet.Fields('SUBMITTER').value;
		sheet.Cells(myRow,3) = recSet.Fields('INVALID_PROBLEM').value;
		sheet.Cells(myRow,4) = recSet.Fields('RESOLUTION').value;
		sheet.Cells(myRow,5) = recSet.Fields('WORKED_DATE').value;
		//sheet.Cells(myRow,6) = recSet.Fields('INVALID_TICKET').value;
		
		myRow++;
		recSet.MoveNext	
	}
	
	//Set recSet = Nothing;
	
	recSet.close
	cObj.close
	
	recSet = null;
	cObj = null;

}


function RunningTotals()
{
	window.grStr2="SELECT * FROM TO_BE_WORKED";
	
	OpenDatabase()
	OpenTable2()
	
	alert(recSet2.RecordCount)
	
	recSet2.close
	cObj.close
	recSet2 = null;
	recSet2 = null;
}


function GetUserRangedStats()
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
	
	ReportFrom = document.frmmain.txt_CheckFrom.value
	ReportTo = document.frmmain.txt_CheckTo.value
    EmpNum = document.frmmain.txt_EmployeeNum.value;
	
	if (ReportFrom != "" || ReportTo != "")
	{
		myConnectString = "SELECT * FROM DefaultTicketStats";
		myConnectString += " WHERE Date BETWEEN #" + ReportFrom + "# AND #" + ReportTo + "# AND EmpNum='" + EmpNum + "'";

		window.grStr3=myConnectString;
		
		UserStats = window.open("pub/templates/ActionsLog.html","actionslog",'width=950,height=350,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');

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

		if(AssignCount != 0)
		{
			LockedCount = AssignCount - TotalCompleted - NeverWorked;
			
			UserStats.document.write('<html>\n');

			UserStats.document.write('	<head>\n');
			UserStats.document.write('		<!-- CSS STYLESHEET -->\n');
			UserStats.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
			UserStats.document.write('		<!-- JAVASCRIPT LINKS -->\n');
			UserStats.document.write('		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n');
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
	}
	else
	{
		alert('A date range is needed! Please choose a from date and to date and try again.')
	}
}

function GetUserHourlyStats()
{
	var x = 0;
	var statsWindow = "";
	var EmpNum = "";
	var ReportFrom = "";
	var ReportTo = "";
	var Display_Hours = new Array('11AM','12PM','1PM','2PM','3PM','4PM','5PM','6PM','7PM','8PM','9PM','10PM','11PM','12AM','1AM');
	var AT11 = new Array();
	var AT12 = new Array();
	var AT13 = new Array();
	var AT14 = new Array();
	var AT15 = new Array();
	var AT16 = new Array();
	var AT17 = new Array();
	var AT18 = new Array();
	var AT19 = new Array();
	var AT20 = new Array();
	var AT21 = new Array();
	var AT22 = new Array();
	var AT23 = new Array();
	var AT24 = new Array();
	var AT0 = new Array();
	var AT1 = new Array();
	
	ReportFrom = document.frmmain.txt_CheckFrom.value
	ReportTo = document.frmmain.txt_CheckTo.value
    EmpNum = document.frmmain.txt_EmployeeNum.value;
	
	//Init. Popup Window
	statsWindow = window.open("pub/templates/Output.html","actionslog",'width=950,height=350,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes');
	
	//Setup SQL Query string
	window.grStr3 = "SELECT * FROM DefaultTicketStats WHERE Date BETWEEN #" + ReportFrom + "# AND #" + ReportTo + "# AND EmpNum='" + EmpNum + "'";
	
	alert(window.grStr3)
	
	OpenDatabase2()	//Open Database
	OpenTable3()	//Open Table
	
	//Count Totals
	while(recSet3.EOF == 0)
	{
		switch(recSet3.Fields('Action').value)
		{
			case 'Assign Ticket':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[0]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[0]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[0]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[0]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[0]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[0]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[0]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[0]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[0]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[0]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[0]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[0]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[0]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[0]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[0]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[0]++
				break
				
			case 'Completed':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[1]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[1]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[1]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[1]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[1]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[1]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[1]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[1]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[1]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[1]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[1]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[1]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[1]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[1]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[1]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[1]++
				break
				
			case 'ToSelf':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[2]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[2]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[2]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[2]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[2]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[2]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[2]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[2]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[2]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[2]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[2]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[2]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[2]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[2]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[2]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[2]++
				break
				
			case 'ToPublic':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[3]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[3]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[3]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[3]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[3]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[3]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[3]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[3]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[3]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[3]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[3]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[3]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[3]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[3]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[3]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[3]++
				break
				
			case 'Reassigned':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[4]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[4]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[4]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[4]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[4]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[4]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[4]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[4]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[4]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[4]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[4]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[4]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[4]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[4]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[4]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[4]++
				break
				
			case 'NeverWorked':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[5]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[5]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[5]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[5]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[5]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[5]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[5]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[5]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[5]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[5]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[5]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[5]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[5]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[5]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[5]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[5]++
				break
			case 'DidCallBack':
				if(recSet3.Fields('RequestTime').value == 11)
					AT11[6]++
				if(recSet3.Fields('RequestTime').value == 12)
					AT12[6]++
				if(recSet3.Fields('RequestTime').value == 13)
					AT13[6]++
				if(recSet3.Fields('RequestTime').value == 14)
					AT14[6]++
				if(recSet3.Fields('RequestTime').value == 15)
					AT15[6]++
				if(recSet3.Fields('RequestTime').value == 16)
					AT16[6]++
				if(recSet3.Fields('RequestTime').value == 17)
					AT17[6]++
				if(recSet3.Fields('RequestTime').value == 18)
					AT18[6]++
				if(recSet3.Fields('RequestTime').value == 19)
					AT19[6]++
				if(recSet3.Fields('RequestTime').value == 20)
					AT20[6]++
				if(recSet3.Fields('RequestTime').value == 21)
					AT21[6]++
				if(recSet3.Fields('RequestTime').value == 22)
					AT22[6]++
				if(recSet3.Fields('RequestTime').value == 23)
					AT23[6]++
				if(recSet3.Fields('RequestTime').value == 0)
					AT0[6]++
				if(recSet3.Fields('RequestTime').value == 1)
					AT1[6]++
				if(recSet3.Fields('RequestTime').value == 2)
					AT2[6]++
				break
			default:
		}
	}		
	
	alert(recSet3.RecordCount)
	
	//Assign Results
	if(recSet3.RecordCount > 0)
	{
		UserStats += '<html>\n';
		UserStats += '	<head>\n';
		UserStats += '		<!-- CSS STYLESHEET -->\n';
		UserStats += '		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n';
		UserStats += '		<!-- JAVASCRIPT LINKS -->\n';
		UserStats += '		<script type="text/javascript" src="T://All In One//default tickets//main_v2.js"></script>\n';
		UserStats += '		<style type="text/css">\n';
		UserStats += '			body {\n';
		UserStats += '				/*background-color: #D0D0D0;*/\n';
		UserStats += '				background-image:url(pub/images/backgrounds/page_bg.jpg;\n';
		UserStats += '				background-repeat: repeat;\n';
		UserStats += '				border-style: none;\n';
		UserStats += '			}\n';
		UserStats += '		</style>\n';
		UserStats += '	</head>\n';
		UserStats += '	<body>\n';
		UserStats += '		<form name="frmagentsstats">';
		UserStats += '			<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n';
		UserStats += '			<tr style="color:#ffffff;" class="text_header">\n';
		UserStats += '				<td>&nbsp;'+ EmpNum +'</td>\n';
		UserStats += '				<td style="text-align:right;"><label for="txt_CloseWindow" accesskey="c"></label><input id="txt_CloseWindow" type="button" value="Close" OnClick="window.close()" class="text_header text_normal"></td>';
		UserStats += '			</tr>\n';
		UserStats += '			<tr>\n';
		UserStats += '				<td>';
		
		UserStats += '					<table width="100%" cellpadding="5" cellspacing="0" border="0" style="text-align:left;" class="tlb_outer">\n';
		UserStats += '						<tr class="text_subheader">\n';
		UserStats += '							<td>Time</td>';
		UserStats += '							<td>Assign Ticket</td>';
		UserStats += '							<td>Resolved</td>';
		UserStats += '							<td>Callbacks Self</td>';
		UserStats += '							<td>Callbacks Public</td>';
		UserStats += '							<td>Reassigned</td>';
		UserStats += '							<td>Never Worked</td>';
		UserStats += '							<td>Callback Done</td>';
		UserStats += '							<td>Total Worked</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[0] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[0] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[0] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[1] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[1] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[1] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[2] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[2] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[2] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[3] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[3] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[3] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[4] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[4] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[4] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[5] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[5] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[5] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[6] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[6] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[6] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[7] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[7] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[7] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[8] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[8] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[8] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[9] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[9] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[9] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[10] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[10] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[10] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[11] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[11] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[11] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[12] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[12] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[12] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[13] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[13] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[13] +'</td>';
		UserStats += '						</tr>\n';
		
		UserStats += '						<tr class="text_normal">\n';
		UserStats += '							<td>' + Display_Hours[14] + '</td>';
		UserStats += '							<td>&nbsp;'+ AT11[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT12[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT13[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT14[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT15[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT16[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT17[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT18[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT19[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT20[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT21[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT22[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT23[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT24[14] +'</td>';
		UserStats += '							<td>&nbsp;'+ AT0[14] +'</td>';
		UserStats += '						</tr>\n';
		UserStats += '					</table>\n';
		
		UserStats += '				</td>\n';
		UserStats += '			</tr>\n';
		UserStats += '			</table>\n';
		UserStats += '		</form>\n';
		UserStats += '	</body>\n';
		UserStats += '</html>';
		
		alert(UserStats)
	}
	
	recSet3.close	//Close Table
	cObj2.close		//Close Database
	
	x = null;
	
	statsWindow = null;
	
	EmpNum = null;
	ReportFrom = null;
	ReportTo = null;
	
	Display_Hours = null;
	
	AT11 = null;
	AT12 = null;
	AT13 = null;
	AT14 = null;
	AT15 = null;
	AT16 = null;
	AT17 = null;
	AT18 = null;
	AT19 = null;
	AT20 = null;
	AT21 = null;
	AT22 = null;
	AT23 = null;
	AT24 = null;
	AT0 = null;
	AT1 = null;
}