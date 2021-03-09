Goal: Create weekly reports at the click of a button

Requirements:
	Needs to pull relevant users and discard irrelevant ones. 
	Needs to be able to be automatically shared
	Needs to end up in a google sheet with appropriate formatting. 
	Needs to account for student situations and plan b
	Needs to automatically calculate averages, %d, etc. 


Step 1: Gather variables and put them into files to be used to help with analysis and report generation.  

Create a configuration file and use config_settings() to put the data from csv into a dict. 
This file is called ‘data.csv’ and the columns provide information that we may want to use later. We can take this information into the python script and use it to pull information from the JSON files we see. 
For instance: The csv header values are put into a dictionary as a key and the values in the column are put as a list into that key. So for the column ‘fieldnames’ the dictionary would look like this. {‘fieldnames’ : [value1, value2, value 3, etc..] }. So when we want to pull specific fields from a report we put it into this file ahead of time so that report raw automatically integrates it with all the other fields it needs to put without having to change report raw’s script. 
For our uses this file holds two pieces of important information:
Fieldnames for the report_raw() function
There are many values that report raw can pull from a user usage report and we only want a small portion of those values. A list of those values can be found here. 
The variables in this list need to be followed with extreme scrutiny. Follow this webpage from google on what variables we can print. https://developers.google.com/admin-sdk/reports/v1/appendix/usage/user
Plan B student names
In calculating active days plan b students who come in the building twice a week need to be accounted for. This is a list of students and the days onto which they come into the building. 
Pull User Data from google directory
Function create_class_list() pulls information directly from OU’s in google to get more user information than is available directly from a user usage report. 
Pulls grade and teacher from OU structure. Ex. \students\grade\teacher
Importantly also tells us if the user is suspended which is important if a user is suspended but not taken out of the school
		


Step 2:  Generate report
Create Date() class
This class will help with date generation for the reports. The user usage report requires a specific date to generate a report. In lieu of a date it just generates a report on the most recent date possible. The Date() class helps with this by creating usable dates that google can use to generate reports. 
When initialized Date() assigns today’s date to itself, yesterday, and 1 week from today. It does this by using the datetime module present in python 3.7. 
It then has several built in functions that make it possible for the Date() module to understand when the closest week for a report is possible. It can take Google up to 3 days to fully transmit all data into the report and thus to generate a full week of reports it will mean the earliest available date would be the monday of the following week. Thus week_report uses Date() to calculate the nearest monday(where a full report is available) regardless of the day of the week. For instance, if you wanted a report for the week of January 4th 2021 on January 10th  it would not be possible because the whole has not been uploaded. If you ran the current build(2-19-21) it would pull on the last week in December of 2020. Essentially Date() allows the program to be one press rather than requiring the user to input a valid date range. 
Authenticate with Google
Google_authenticate(): This function uses a service account created in google and uses a credentials.json file to pull information for that account**. 
** It would be good to look at ways of securing this information.** 
Pickle file is something that was used in the sample script given by google and for this purpose is a secure way of seeing if the user has logged in recently(or so I think. More research may need to be done)
Create reports and put them into an anaylzed_data() dictionary
# this is a large section that will require major explanations. 
#do not take this as a complete document in regards to understanding the whole scope of week_report()
week _report() is a function that generates 6 reports (sunday-friday) and compares each of them. The basic idea is to check if values change between days of the week and if they do to mark the person present for that specific day and application. week_report() uses the Date() class to help with this. Refer to Date() for more information regarding this topic. 
Generate reports
report_raw 
report_raw (RR) is the fundamental tool that generates reports from our gsuite side. It takes 5 arguments: 
The date in YYYY/MM/DD format
unused  potential for userkey used to specify which groups of users
Unused potential for parameters used to pass to report request
Maxresults in integer format
RR uses many other formulas inside of it to make what it does possible. First it uses google_authenticate() to generate the credentials for the report. Config_settings() to get the fields that we want to take from the report, and also Date() to help with date calculations. Refer to them in the program and in the above for their uses. 
RR generates first a user_usage report which comes in the form of a JSON file. Essentially just a bunch of dictionaries within dictionaries. Each user is given an entry in the report with email and name given but more importantly their respective data. RR takes this data and pulls PII from each entry and then the data that goes along with this. RR does make sure to eliminate certain people                                                  from the data list. Specifically: teachers. It puts it into a dictionary where the keys are the emails of the users and the values are another dictionary where the keys are the field names for each data point and values are the values to each fieldname. An example would be:
{nsutdent@raleighoakcharter.org : {Fieldname: data, etc.} }

RR finally exports all user data in the form of the above mentioned dictionary. 

 	
report_raw() is used to generate, as the name suggests, raw reports but the data needs to be analyzed to other days to be able to see trends. That is where report_to_report()  and report_to_data() come into play.
Report_to_report()
This is the initializing data analysis function. Report_to _report(RtR) in the context of week_report() takes Sunday’s (report_1) and Monday’s(report_2) report and compares the two. It generates a new dictionary that’s structure is used for the rest of the analysis process. This dictionary is similar to the RR dict in that it’s first index of keys is user emails and then related data but different in that the data is now more analyzed than before and initialized based on different factors. Example : 



There are many points compared but to save time essentially it looks to see what the report says about user usage time and then makes comparisons between both days. From those comparisons it will add to or subtract from active days or missed days (plan b students are not being accounted for in this part of the code see further_analysis(). 
There is an interesting trick used to add data to the dictionary. One of the issues that came up was how to add each date in a way that makes sense in a work week. For instance how does the code know to add to the ‘monday’ category. RtR requires an argument specifying the day of the week. This is every letter up to the word in each weekday so for Monday the day of the week would be mon because mon + day = monday. Tuesday would be tues and so on and so forth. This allows the report to add data to the correct place. ** future development idea** : instead of asking for the user to do this just simply take the date from the report and calculate what day of the week it is using the datetime module. Skipping the user input
Lastly the above dictionary is exported so that it can be continued to be added to. 						

Report_to_data()
Report_to_data(RtD) This function is a continuation of the RtR function. Much like RtR, RtD takes information from two reports and compares them but it does it using the analyzed_data() from other analysis. RtR is meant to create the analyzed_data() dictionary whereas RtD is meant to take the analyzed_data() dictionary and just add data elements to it. **development idea** try to integrate both RtR and RtD into on single formula that looks to see if analyzed_data() exists if it does not then it will create it. 
RtD works exactly the same as RtD but it does not initialize the dict. 
RtR exports analyzed data once it is finished which can then be used again with other RtRs 
Incomplete analyzed_data() is created by making use of the above two functions. First Monday and Sunday use the RtR function to create analyzed_data() and then tuesday - friday use RtD each adding onto analyzed_data() until we have an incomplete analyzed_data(). 
Further_data_anylasis() this is the final function of week_report() it receives the incomplete analyzed_data() dictionary. From there it adds a new key called ‘data’ where the rest of analysis goes. There are about 160 and growing calculations (will not cover all of them) but simply it calculates out who is plan b, creates averages based on grade level and school level, and creates trend data points. 
Finally week_report() exports  the complete analyzed_data()
Format data into a google spreadsheet:
This is also extremely complicated and annoying because it is a combination of several modules like gspread, gspread_formatting, and the google api module. All of these make it very difficult to make changes to google spreadsheets but it did eventually work out. It is recommended that you carefully study and look through all of the API documentation for sheets that you can. 
