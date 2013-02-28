outlook
=======

Outlook auto archiving script. (I strongly suggest you use cached mode with this otherwise it will take hours to run and hang every time the event fires)

Tested with Outlook 2010.

Open Outlook, add the Developer ribbon, open Visual Basic and copy/paste the contens or archive.vba into "ThisOutlookSession"
This archiver expects that you have a data file added, named for each year of email "2012", "2013". 
It could be made to create the pst and link it, I just haven't done it since I create the folders myself etc.

Once that is done, close and re-open outlook.
Create a reminder called "MOVEOLDEMAILREMINDER", set it to right now, make sure you set it to set a reminder. Save it.
When it fires, the date will be adjusted to the same time, the next day and all old email will be moved and the reminder dismissed without interaction.
