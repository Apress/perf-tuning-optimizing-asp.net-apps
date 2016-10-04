README for Chapter 7 Web Application Stress Testing and Monitoring

The AspNetChap7ACTTest project provides test samples to illustrate the various features of the Application Center Testing tool.

Here is a summary of each test script:

Project: AspNetChap7ACTTest 
---------------------
Purpose:  	Illustrates the features of the Application Center Testing tool.  
	Each test is designed to run against the Duwamish 7.0 sample site included with .NET.
	To install the Duwamish source files and database, run C:\Program Files\Microsoft Visual Studio .NET\Enterprise Samples\Duwamish 7.0 VB\Duwamish.msi.

Test Script			Purpose
------------------------------	------------------------------
Login_Select_Checkout_1	A robust browser recorded session which goes through the process of 
			logging into the application, selecting several items, and proceeding through checkout.

Category_Fiction_History	A more focused browser recorded session which requests the Fiction and History category pages.
			Several test runs were performed with this test script, with a wide range of simultaneous connnection settings.
			Results are saved as part of the project.
			Much of the performance testing samples and analysis focus on tests run with this test script.

BrowserTest		A scripted dynamic test which illustrates how to set up the header information using the Test Object Model
			to simulate a request from a specific browser (IE 6.0)

PageTest			A scripted dynamic test which illustrates how to use generate a set of requests from the test script
			using a .dat file to store the request paths.

File Name
------------------------------	------------------------------
ScheduleTaskRunScript1.vbs	This script illustrates how to run ACT from a script.  
			The script can be set up and run unatteneded as a scheduled task.

RequestPaths.dat		Stores the request paths for the PageTest script.