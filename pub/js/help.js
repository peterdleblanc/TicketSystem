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

function openHelp()
{
	var sHelpContents;
	var iRecCount;
	
	window.grStr="SELECT * FROM tblTOC_Contents WHERE ControlEnabled = '1' ORDER BY sort_id ASC";
	
	sHelpContents = window.open("pub/templates/HelpContents.html","HelpContents",'width=800,height=450,toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes\n');
	
	OpenDatabase()
	OpenTable()
	
	iRecCount = 1;
	recSet.MoveFirst;
	
	sHelpContents.document.write('<html>\n');
	sHelpContents.document.write('	<head>\n');
	sHelpContents.document.write('		<!-- CSS STYLESHEET -->\n');
	sHelpContents.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
	sHelpContents.document.write('		<!-- JAVASCRIPT LINKS -->\n');
	sHelpContents.document.write('		<script type="text/javascript" src="pub\\js\\help.js"></script>\n');
	sHelpContents.document.write('		<style type="text/css">\n');
	sHelpContents.document.write('			body {\n');
	sHelpContents.document.write('				background-image:url(pub/images/backgrounds/page_bg.jpg);\n');
	sHelpContents.document.write('				background-repeat: repeat;\n');
	sHelpContents.document.write('				border-style: none;\n');
	sHelpContents.document.write('			}\n');
	sHelpContents.document.write('		</style>\n');
	sHelpContents.document.write('	</head>\n');
	sHelpContents.document.write('	<body>\n');
	sHelpContents.document.write('		<form name="frmHelpContents">\n');
	sHelpContents.document.write('			<table width="100%" cellspacing="0" cellpadding="0">\n');
	sHelpContents.document.write('				<tr>\n');
	sHelpContents.document.write('					<td>\n');
	
	sHelpContents.document.write('					<table width="100%" cellspacing="0" cellpadding="0" class="tlb_outer">\n');
	sHelpContents.document.write('						<tr>\n');
	sHelpContents.document.write('							<td>\n');
	sHelpContents.document.write('								<table width="100%" cellspacing="0" cellpadding="0" class="text_header">\n');
	sHelpContents.document.write('									<td>&nbsp;Online Manual</td>\n');
	sHelpContents.document.write('									<td style="text-align:right;"><label for="cmd_Close" accesskey="c"><input id="cmd_Close" type="button" value="Close" class="text_subheader text_normal text_buttons" onClick="window.close()"></label></td>\n');
	sHelpContents.document.write('								</table>');
	sHelpContents.document.write('							</td>\n');
	sHelpContents.document.write('						</tr>\n');
	sHelpContents.document.write('						<tr>\n');
	sHelpContents.document.write('							<td>\n');
	
	/*
	sHelpContents.document.write('								<!-- Table of Contents -->\n');
	sHelpContents.document.write('								<table width="100%" cellspacing="0" cellpadding="0" class="tlb_outer">\n');
	sHelpContents.document.write('									<tr class="text_subheader">\n');
	sHelpContents.document.write('										<td>&nbsp;Table of Contents</td>\n');
	sHelpContents.document.write('									</tr>\n');
	sHelpContents.document.write('									<tr class="text_normal">\n');
	sHelpContents.document.write('										<td>\n');
	sHelpContents.document.write('											<table width="100%" cellspacing="0" cellpadding="0" class="text_normal">\n');

	recSet.MoveFirst;
	
	while(!recSet.eof) {
		sHelpContents.document.write('											<tr>\n');
		sHelpContents.document.write('												<td>' + iRecCount + '. <a href="#" onClick="OpenChapter(' + iRecCount + ')">' + recSet.Fields("title").value + '</a></td>\n');
		sHelpContents.document.write('											</tr>\n');
		
		iRecCount++;
		recSet.MoveNext
	}
	
	sHelpContents.document.write('											</table>\n');
	sHelpContents.document.write('										</td>\n');
	sHelpContents.document.write('									</tr>\n');
	sHelpContents.document.write('								</table>\n');
	sHelpContents.document.write('								<!-- Table of Contents -->\n');
	*/
	sHelpContents.document.write('							</td>\n');
	sHelpContents.document.write('						</tr>\n');
	
	iRecCount = 1;
	recSet.MoveFirst;
	
	sHelpContents.document.write('						<!-- Contents -->\n');
	sHelpContents.document.write('						<tr>\n');
	sHelpContents.document.write('							<td>\n');
	sHelpContents.document.write('								<table width="100%" cellspacing="0" cellpadding="1" class="tlb_outer" style="color:#ffffff;">\n');
	
	while(!recSet.eof) {
		sHelpContents.document.write('									<!-- Chapter ' + iRecCount + ':' + recSet.Fields("title").value + ' -->\n');
		
		
		
		
		sHelpContents.document.write('									<tr class="text_subheader">\n');
		sHelpContents.document.write('										<td>Chapter ' + iRecCount + ': ' + recSet.Fields("title").value + '</td>\n');
		sHelpContents.document.write('									</tr>\n');
		
		sHelpContents.document.write('									<tr class="text_toc-normal">\n');
		sHelpContents.document.write('										<td>'+recSet.Fields("text").value+'</td>\n');
		sHelpContents.document.write('									</tr>\n');
		
		iRecCount++;
		recSet.MoveNext
	}
	
	sHelpContents.document.write('									</table>\n');
	sHelpContents.document.write('								</td>\n');
	sHelpContents.document.write('						</tr>\n');
	sHelpContents.document.write('						<!-- Contents -->\n');
	
	sHelpContents.document.write('					</table>\n');
	
	sHelpContents.document.write('					</td>\n');
	sHelpContents.document.write('				</tr>\n');
	sHelpContents.document.write('			</table>\n');
	
	sHelpContents.document.write('		</form>\n');
	sHelpContents.document.write('	</body>\n');
	sHelpContents.document.write('</html>\n');
	
	recSet.close
	cObj.close
}

function OpenChapter(chapter_num)
{
	var sHelpContents;
	var iRecCount;
	
	window.grStr="SELECT * FROM tblTOC_Contents WHERE id = '" + chapter_num + "'";
	
	sHelpContents = window.open("pub/templates/Help.html","HelpContents",'width=800,height=450,toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,copyhistory=no, resizable=yes\n');
	
	OpenDatabase()
	OpenTable()
	
	iRecCount = 1;
	recSet.MoveFirst;
	
	sHelpContents.document.write('<html>\n');
	sHelpContents.document.write('	<head>\n');
	sHelpContents.document.write('		<!-- CSS STYLESHEET -->\n');
	sHelpContents.document.write('		<link rel="stylesheet" media="screen" href="styles.css" type="text/css">\n');
	sHelpContents.document.write('		<!-- JAVASCRIPT LINKS -->\n');
	sHelpContents.document.write('		<script type="text/javascript" src="pub\\js\\help.js"></script>\n');
	sHelpContents.document.write('		<style type="text/css">\n');
	sHelpContents.document.write('			body {\n');
	sHelpContents.document.write('				background-image:url(pub/images/backgrounds/page_bg.jpg);\n');
	sHelpContents.document.write('				background-repeat: repeat;\n');
	sHelpContents.document.write('				border-style: none;\n');
	sHelpContents.document.write('			}\n');
	sHelpContents.document.write('		</style>\n');
	sHelpContents.document.write('	</head>\n');
	sHelpContents.document.write('	<body>\n');
	sHelpContents.document.write('		<form name="frmHelpContents">\n');
	sHelpContents.document.write('			<table width="100%" cellspacing="0" cellpadding="0">\n');
	sHelpContents.document.write('				<tr>\n');
	sHelpContents.document.write('					<td>\n');
	
	sHelpContents.document.write('					<table width="100%" cellspacing="0" cellpadding="0" class="tlb_outer">\n');
	sHelpContents.document.write('						<tr>\n');
	sHelpContents.document.write('							<td>\n');
	sHelpContents.document.write('								<table width="100%" cellspacing="0" cellpadding="0" class="text_header">\n');
	sHelpContents.document.write('									<td>&nbsp;Online Manual</td>\n');
	sHelpContents.document.write('									<td style="text-align:right;"><label for="cmd_TableofContents" accesskey="t"><input id="cmd_TableofContents" type="button" value="Table Of Contents" class="text_subheader text_normal text_buttons" onClick="OpenTOC()"></label><label for="cmd_Close" accesskey="c"><input id="cmd_Close" type="button" value="Close" class="text_subheader text_normal text_buttons" onClick="window.close()"></label></td>\n');
	sHelpContents.document.write('								</table>');
	sHelpContents.document.write('							</td>\n');
	sHelpContents.document.write('						</tr>\n');
	
	iRecCount = 1;
	recSet.MoveFirst;
	
	sHelpContents.document.write('						<!-- Contents -->\n');
	sHelpContents.document.write('						<tr>\n');
	sHelpContents.document.write('							<td>\n');
	sHelpContents.document.write('								<table width="100%" cellspacing="0" cellpadding="1" class="tlb_outer">\n');
	
	while(!recSet.eof) {
		sHelpContents.document.write('									<!-- Chapter ' + iRecCount + ':' + recSet.Fields("title").value + ' -->\n');
		
		
		
		
		sHelpContents.document.write('									<tr class="text_subheader">\n');
		sHelpContents.document.write('										<td>Chapter ' + iRecCount + ': ' + recSet.Fields("title").value + '</td>\n');
		sHelpContents.document.write('									</tr>\n');
		
		sHelpContents.document.write('									<tr class="text_toc-normal">\n');
		sHelpContents.document.write('										<td>'+recSet.Fields("text").value+'</td>\n');
		sHelpContents.document.write('									</tr>\n');
		
		iRecCount++;
		recSet.MoveNext
	}
	
	sHelpContents.document.write('									</table>\n');
	sHelpContents.document.write('								</td>\n');
	sHelpContents.document.write('						</tr>\n');
	sHelpContents.document.write('						<!-- Contents -->\n');
	
	sHelpContents.document.write('					</table>\n');
	
	sHelpContents.document.write('					</td>\n');
	sHelpContents.document.write('				</tr>\n');
	sHelpContents.document.write('			</table>\n');
	
	sHelpContents.document.write('		</form>\n');
	sHelpContents.document.write('	</body>\n');
	sHelpContents.document.write('</html>\n');
	
	recSet.close
	cObj.close
}