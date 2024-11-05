Instructions to run:
1, Double click on LoggingTool.bat to start.
1.1, If you get a Windows pop up a saying Windows Protected Your PC, click on More Info, then click on Run Anyway.
2, You will get a pop up window asking how long you would like the logging session to run.
3, Clicking on your preferred option to start logging.
4, Once done, a .csv file will be saved and a window will open to show the saved location.
5, Double click on ReportTool.bat and you will get to locate the .csv file.
6, Select the .csv file, the ReportTool will start processing and spit out a .html file to show the visualized report in line charts.

Understand the report:
1, Processes are colored in RED, YELLOW and GREEN based on the highest processor time. RED is > 20%, YELLOW is between 5% to 20% and GREEN is lower than 5%
2, You may see processor time being bigger than 100%, this is due to the tool logs usage based on CPU cores. For example, if you see a process with 200% usage, it could mean 2 CPU cores are being fully utilized.

Addition:
1, PS folder contains the actual PowerShell script.
2, .bat file just to call the scripts and apply -ExecutionPolicy Bypass