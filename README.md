# PPT-to-OBS-cmd
A PowerPoint plugin that sends commands to OBS on slide show page changes.

This add-on script allows you to use PowerPoint speaker notes to control OBS using OBSCommand.
Learn more about OBSCommand here: https://github.com/REALDRAGNET/OBSCommand

To add this to your PowerPoint application, go to the Releases tab and download the .ppam file.
You can use this add-on by opening the file directly from the file explorer.
OR, you can open PowerPoint, go to **Options** > **Add-ins** and select the drop-down menu next to **Manage** and select **PowerPoint Add-ins**. Then click **Go**. Click **Add New**. Then add the .ppam file and enable the add-on.

The script works by mapping speaker note strings to batch file names. If the first sequence of characters (before the first space delimiter) matches a string in this map, then the corresponding batch file is called and the commands for OBSCommand are run.
The batch files for OBSCommand should be created at "C:\Program Files\OBSCommand\scenes\".

The SceneMap.csv file allows you to customize speaker note strings in Column A and map them to batch file names in Column B of the CSV. This CSV file should be made at "C:\Program Files\PPT-to-OBS-cmd\SceneMap.csv".
Example SceneMap.csv:
CAM+SCREEN,OBS1  
SCREENONLY,OBS2  
CAMONLY,OBS3

For instance, we place "CAM+SCREEN" as the first sequence of characters in the speaker notes on slide 2. Then the add-on script runs when we advance to slide 2, detecting "CAM+SCREEN" and executing the batch file containing commands for OBSCommand in OBS1.bat, located in "C:\Program Files\OBSCommand\scenes\".

