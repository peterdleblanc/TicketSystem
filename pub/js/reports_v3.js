	function Database_Class(type)
	{
		this.Path = "T:\\All In One\\default_tickets\\pub\\databases\\";
		this.iStr = "";
		this.DBName = "";
		this.QryString = "";
		this.QryString2 = "";
		this.adOpenKeyset = '1';
		this.adLockOptimistic = '2';
		this.adCmdText = '1';
		this.adUseClient = '3';

		this.cObj = new ActiveXObject("ADODB.Connection");
		this.recSet = new ActiveXObject("ADODB.Recordset");
		this.recSetCSV = new ActiveXObject("ADODB.Recordset");
		
		this.OpenTable = function()
		{
			this.iStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.Path + this.DBName + ";Persist Security Info=False";
			this.cObj.Open(this.iStr)
			this.recSet.Open(this.QryString,this.cObj,this.adOpenKeyset,this.adLockOptimistic,this.adCmdText)
		}
		
		this.CloseTable = function()
		{
			this.recSet.close
			this.cObj.close
			this.recSet = null;
			this.cObj = null;
		}
		
		this.OpenCSVFile = function()
		{
			this.CSVStr='Provider=MSDASQL; Driver={Microsoft Text Driver (*.txt; *.csv)}; DBQ=T:\\All In One\\default_tickets\\';
			cObj.Open(this.CSVStr)
			
			recSetCSV.Open(QryString2,cObjCSV,adOpenKeyset,adLockOptimistic,adCmdText)
		}
		
		this.CloseCSVFile = function()
		{
			cObj.close
			
			this.recSetCSV = null;
			this.cObj = null;
			
		}
	
		this.LockedRecordCount = function()
		{
			this.DBName = "DefaultTickets_revised.mdb"
			this.iStr = "SELECT * FROM TO_BE_WORKED WHERE BEING_WORK_BY IS NOT NULL AND BEING_WORK_BY <> 'z'";
			
			this.OpenTable()
			
			alert(this.recSet.RecordCount)
			
			this.CloseTable()
		}
		
		this.ImportCSVFile = function()
		{
				for(i=0;i<this.recSetCSV.RecordCount;i++)
				{
					this.recSet.AddNew
						this.recSet.Fields('ASSIGNED_DATE').value=this.recSetCSV.Fields(0).value  		//ASSIGNED_DATE
						this.recSet.Fields('ASSIGNED_GROUP').value=this.recSetCSV.Fields(1).value			//ASSIGNED_GROUP
						this.recSet.Fields('ASSIGNED_TO').value=this.recSetCSV.Fields(2).value			//ASSIGNED_TO
						this.recSet.Fields('AUTO_CORRELATE').value=this.recSetCSV.Fields(3).value			//AUTO_CORRELATE
						this.recSet.Fields('CATEGORY').value=this.recSetCSV.Fields(4).value			//CATEGORY
						this.recSet.Fields('CORRELATION_TYPE').value=this.recSetCSV.Fields(5).value			//CORRELATION_TYPE
						this.recSet.Fields('CORRELATION_VALUE_1').value=this.recSetCSV.Fields(6).value			//CORRELATION_VALUE_1
						this.recSet.Fields('CORRELATION_VALUE_2').value=this.recSetCSV.Fields(7).value			//CORRELATION_VALUE_2
						this.recSet.Fields('CREATE_DATE').value=this.recSetCSV.Fields(8).value			//CREATE_DATE
						this.recSet.Fields('DEPARTMENT').value=this.recSetCSV.Fields(9).value			//DEPARTMENT
						this.recSet.Fields('ETA').value=this.recSetCSV.Fields(10).value			//ETA
						this.recSet.Fields('INDICATOR_CITY').value=this.recSetCSV.Fields(11).value			//INDICATOR_CITY
						this.recSet.Fields('INDICATOR_STATE').value=this.recSetCSV.Fields(12).value			//INDICATOR_STATE
						this.recSet.Fields('MultiRegion Affected Names').value=this.recSetCSV.Fields(13).value			//MultiRegion Affected Names
						this.recSet.Fields('MultiRegion Affected').value=this.recSetCSV.Fields(14).value			//MultiRegion Affected
						this.recSet.Fields('MultiRegion Flag').value=this.recSetCSV.Fields(15).value			//MultiRegion Flag
						this.recSet.Fields('NETWORK_ELEMENT').value=this.recSetCSV.Fields(16).value			//NETWORK_ELEMENT
						this.recSet.Fields('NEXT_ACTION').value=this.recSetCSV.Fields(17).value			//NEXT_ACTION
						this.recSet.Fields('NIU ID').value=this.recSetCSV.Fields(18).value			//NIU ID
						this.recSet.Fields('Node').value=this.recSetCSV.Fields(19).value			//Node
						this.recSet.Fields('OFFICE_NAME').value=this.recSetCSV.Fields(20).value			//OFFICE_NAME
						this.recSet.Fields('PARENT_TICKET').value=this.recSetCSV.Fields(21).value			//PARENT_TICKET
						this.recSet.Fields('PRIMARY_PHONE').value=this.recSetCSV.Fields(22).value			//PRIMARY_PHONE
						this.recSet.Fields('PRIORITY').value=this.recSetCSV.Fields(23).value			//PRIORITY
						this.recSet.Fields('PROBLEM_CODE').value=this.recSetCSV.Fields(24).value			//PROBLEM_CODE
						this.recSet.Fields('PROBLEM_SUMMARY').value=this.recSetCSV.Fields(25).value			//PROBLEM_SUMMARY
						this.recSet.Fields('PRODUCT').value=this.recSetCSV.Fields(26).value			//PRODUCT
						this.recSet.Fields('PURGE_DATE').value=this.recSetCSV.Fields(27).value			//PURGE_DATE
						this.recSet.Fields('QUEUE_ID').value=this.recSetCSV.Fields(28).value			//QUEUE_ID
						this.recSet.Fields('QUEUE_NAME').value=this.recSetCSV.Fields(29).value			//QUEUE_NAME
						this.recSet.Fields('REGION_ID').value=this.recSetCSV.Fields(30).value			//REGION_ID
						this.recSet.Fields('REGION_NAME').value=this.recSetCSV.Fields(31).value			//REGION_NAME
						this.recSet.Fields('SA_ID').value=this.recSetCSV.Fields(32).value			//SA_ID
						this.recSet.Fields('SERVICE_AFFECT').value=this.recSetCSV.Fields(33).value			//SERVICE_AFFECT
						this.recSet.Fields('Severity').value=this.recSetCSV.Fields(34).value			//Severity
						this.recSet.Fields('STATUS').value=this.recSetCSV.Fields(35).value			//STATUS
						this.recSet.Fields('SUBMITTER').value=this.recSetCSV.Fields(36).value			//SUBMITTER
						this.recSet.Fields('SYSTEM_ID').value=this.recSetCSV.Fields(37).value			//SYSTEM_ID
						this.recSet.Fields('SYSTEM_NAME').value=this.recSetCSV.Fields(38).value			//SYSTEM_NAME
						this.recSet.Fields('TECH_ARRIVAL_TIME').value=this.recSetCSV.Fields(39).value			//TECH_ARRIVAL_TIME
						this.recSet.Fields('TECHNICIAN_NAME').value=this.recSetCSV.Fields(40).value			//TECHNICIAN_NAME
						this.recSet.Fields('TECHNICIAN_NUMBER').value=this.recSetCSV.Fields(41).value			//TECHNICIAN_NUMBER
						this.recSet.Fields('TICKET_ID').value=this.recSetCSV.Fields(42).value			//TICKET_ID
						this.recSet.Fields('TTR_SLA').value=this.recSetCSV.Fields(43).value			//TTR_SLA
						this.recSet.Fields('TTR_START').value=this.recSetCSV.Fields(44).value			//TTR_START
						this.recSet.Fields('WF-Scratch1').value=this.recSetCSV.Fields(45).value			//WF-Scratch1
					this.recSet.update
					this.recSetCSV.MoveNext
					this.recSet.MoveNext
				}

		}
	}
	
	