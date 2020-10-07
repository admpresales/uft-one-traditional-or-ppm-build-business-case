'===========================================================
'20201007 - DJ: Initial creation
'===========================================================

'===========================================================
'Function to search for the PPM proposal in the appropriate status
'===========================================================
Function PPMProposalSearch (CurrentStatus, NextAction)
	'===========================================================================================
	'BP:  Click the Search menu item
	'===========================================================================================
	Browser("Search Requests").Page("Dashboard - Portfolio").Link("SEARCH").Click
	
	'===========================================================================================
	'BP:  Click the Requests text
	'===========================================================================================
	Browser("Search Requests").Page("Dashboard - Portfolio").Link("Requests").Click @@ hightlight id_;_10000000_;_script infofile_;_ZIP::ssf1.xml_;_
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter PFM - Proposal into the Request Type field
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").WebEdit("Request Type Field").Set "PFM - Proposal"
	Browser("Search Requests").Page("Search Requests").WebElement("StatusLabel").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter a status of "New" into the Status field
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").WebEdit("Status Field").Set CurrentStatus
	
	'===========================================================================================
	'BP:  Click the Search button (OCR not seeing text, use traditional OR)
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").Link("Search").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the first record returned in the search results
	'===========================================================================================
	DataTable.Value("dtFirstReqID") = Browser("Search Requests").Page("Request Search Results").WebTable("Req #").GetCellData(2,2)
	Browser("Search Requests").Page("Request Search Results").Link("First Request ID").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
End Function

Dim BrowserExecutable, Counter

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon
Set AppContext2=Browser("CreationTime:=1")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Strategic Portfolio link
'===========================================================================================
Browser("Search Requests").Page("Project & Portfolio Management").Image("Strategic Portfolio Image").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Barbara Getty (Business Relationship Manager) link to log in as Barabara Getty
'===========================================================================================
Browser("Search Requests").Page("Portfolio Management").WebArea("Barbara Getty Image").Click
AppContext.Sync																				'Wait for the browser to stop spinning

PPMProposalSearch "New", "Approved"

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Approved").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Enter "US" into the Region field
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebEdit("Region").Set "US"

'===========================================================================================
'BP:  Click the Completed button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Completed").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Select "Innovation" in the Project Class
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebList("Project Class").Select "Innovation"

'===========================================================================================
'BP:  Select "Infrastructure" in the Asset Class
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebList("Asset Class").Select "Infrastructure"

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Approved").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Enter the Expected Start Period as June 2021
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebEdit("Expected Start Period").Set "June " & (Year(Now)+1)

'===========================================================================================
'BP:  Enter the Expected Finish Period as December 2021
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebEdit("Expected Finish Period").Set "December " & (Year(Now)+1)

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Completed button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Completed").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Create button
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebElement("Create").Click
AppContext2.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Create button in the popup window
'===========================================================================================
Browser("Create a Blank Staffing").Page("Create a Blank Staffing").WebButton("button.create").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Select the Staffing Profile button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Staffing Profile").Link("Select the Staffing Profile").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Enter "A/R Billing Upgrade" into the Staffing Profile field
'===========================================================================================
Browser("Create a Blank Staffing").Page("Staffing Profile").Frame("copyPositionsDialogIF").WebEdit("Staffing Profile Text").Set "A/R Billing Upgrade"
Browser("Create a Blank Staffing").Page("Staffing Profile").Frame("copyPositionsDialogIF").WebElement("Staffing Profile Label").Click

'===========================================================================================
'BP:  Click the Import button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Staffing Profile").Frame("copyPositionsDialogIF").Link("Import").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Done text
'===========================================================================================
Browser("Create a Blank Staffing").Page("Staffing Profile").WebElement("Done").Click

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action Button").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
Browser("Search Requests").Page("Req Details").Link("Sign Out Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

