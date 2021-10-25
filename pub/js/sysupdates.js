function CheckforUpdates()
{
	var StatsDB = new Database_class('StatsDB')
	var UpdatesQueue = new GenerateAvailableUpdates()
	var Enddate;
	var tableInsert = "";
	var x;
	//Get Date Values
	CurDate = get_GenerateDate(0);
	Enddate = get_GenerateDate(5);
	//Set DB Values
	StatsDB.Path = 'T:\\All In One\\default_tickets\\pub\\js\\databases\\';
	StatsDB.DBName = 'StatsDB.mdb';
	StatsDB.grStr = "SELECT * FROM SystemUpdates WHERE Between #"+CurDate+"# TO #"+Enddate+"#AND enabled=1";
	//Open Table
	StatsDB.OpenTable()	
	//Populate Table with updates.
	while(StatsDB.recSet.EOF) UpdatesQueue.insert(StatsDB.recSet.Fields('details').value);	
	//Get generated table.
	tableInsert = UpdatesQueue.output('text_header','text_normal');
	//Display Table.
	document.write(tableInsert);
	//Close Table
	StatsDB.CloseTable()
	//Cleanup
	StatsDB = null;
	UpdatesQueue = null;
	CurDate = null;
	tableInsert = null;
	x = null;
}

function get_GenerateDate(num_days)
{
	var generated_date = new Date();
	var sDay = generated_date.getDate() + num_days;
	var sMonth = generated_date.getMonth();
	var sMonth = sMonth + 1;
	var sYear = generated_date.getYear();
	var CurrentDate = sMonth + "/" + sDay + "/" + sYear;
	
	return CurrentDate;
	
	generated_date = null;
	sDay = null;
	sMonth = null;
	sMonth = null;
	sYear = null;
	CurrentDate = null;
}

function GenerateAvailableUpdates()
{
	this.data = new Array();
	this.data_index = 0;
	this.insertstr = "";
	
	this.insert = function(data)
	{
		if(data != "") this.data[this.data_index] = data;
	}
	
	this.output = function(headingtext,normaltext)
	{
		this.insertstr = "";
		
		this.insertstr += '<form name="frmSystemUpdates">';
		this.insertstr += '	<div id="div_SystemUpdates" name="div_SystemUpdates" style="position:absolute; visibility:show; top:250px; z-index:1000;">';		
		this.insertstr += '		<table cellpadding="0" cellspacing="0" border="0" class="'+ normaltext +'"';
		this.insertstr += '			<td class="'+ headingtext +'">';
		this.insertstr += '				<tr>Available System Updates</tr>';
		this.insertstr += '			<td>';
		for(var x=0;x<this.data.length;x++)
		{
			this.insertstr += '			<td>';
			this.insertstr += '				<tr>'+ this.data[x] +'</tr>';
			this.insertstr += '			<td>';
			
			this.data_index++
		}
		this.insertstr += '		</table>';
		this.insertstr += '	</div>';
		this.insertstr += '</form>';
		
		if(this.insertstr.length > 0) return this.insertstr;
	}
}