# Hazmat/Code of Conduct Tracker
Here is a script that was created with Python, VBA and a slack webhook 

## Code Summary
1. Harvester.py | Python
2. HazmatCoCTracker | VBA
3. DOIt | VBA
4. theChef_2 | Python
5. Webhook | Slack Webhook

### Automation 
1. ScheduleRun | Python
2. RunHazmat | Python

# Workflow Summary 
### Harvester.py
In the Harvester file it will take in a CSV (Comma seperated Value) file of associates from a dialog box.
After it gets the information from the csv file it will then use pandas to read the information within the file 
and seperate it all out into the table below 

![image](https://github.com/DonovanKilpela/Hazmat-CodeOfConduct-Tracker/assets/161644962/f7bbc852-53f8-41fd-a441-c63f6d6c5d8d)

### HazmatCoCTraker VBA
In the HazmatCoCTracker VBA code it will run all of the python files in a specfic order so the program can run correctly 

### DoIt VBA
In the DoIt VBA code it will check the attendance of who is onsite and will pull the information into the table tell us who is here and clocked in 

### TheChef_2.py
In TheChef_2 it will take the information in from the table of which associates are onsite. It will then break up the information into groups.
The groups that it puts them into are the different jobs that you can do at our site. After everything is put into their seperate groups we used 
the tabulate libray to format the payload that we send to the webhook. 

### ScheduleRun.py
The ScheduleRun file will set specfic times for the scripts to run so we can make sure we are getting as many of the Flex associates as possible. 
Flex associates can pick up a variety of shifts that won't always start at the same time each day so catching them was a issue before we made this. 

### Slack WebHook 
Here is how the webhook looks when it is sent to slack. It isn't the whole table we left out names and also logins 

![image](https://github.com/DonovanKilpela/Hazmat-CodeOfConduct-Tracker/assets/161644962/1cdc9661-5aa3-4d7a-a76c-3003ef8cdac7)


# My Contribution 
1. I created the Harvester, which will use pandas to get the information pasted into the table 
2. I created the Slack webHook that is connected to theChef_2
3. I deep dived into how to make the table through the webhook look neat or organized
4. Created the scheduled messages and buttons through slack to let the team know to add the csv file into the sheet

# Outcomes
Since making this program it lead to a 3.03% MoM increase in complaince for our building on Hazmat and CoC. 
It has made getting the Flex associates faster and helped to increase our compliance dramatically. 
Lastly it made our process faster, before when doing this task we had to manually sort through a huge list of associates, now we just do it with a click of a button. 

# Future Updates 
Right now we are working on getting to fully automate our process so all we have to do is send the messages and get the assocaites trained. To do that
we are hoping to get access to the database so we can pull the information from there instead of through the current means we are now.
