# hivemonitor

<img src="sensor-frame-front.jpeg" alt="Sensor in a frame" width="600px" height="800px">

## Steps and code to build a temperature and humidity monitoring system with SwitchBot sensors.

Introduction is in "hivemonitor.pdf" - it describes the components used (with Amazon links valid in Nov 2024) and how sensors can be fitted into a hive.

Steps listed in the introduction include:
  - procuring sensors and a "mini-hub" (directly from switch-bot.com or via Amazon)
  - building frames for sensors (a DIY job)
  - fitting frames with sensors into your hive(s)
  - installing SwitchBot app
    Important: _When an app asks if sensor data should be stored in the cloud say "Yes". It will be needed for automating the system._
  - connecting the "mini-hub" to Wifi and to the Internet (follow instructions which come with the hub)
  - establishing communication between the app and the sensors (follow instructions which come with the sensors)
  - verifying the graphs in the app work well (you can give hives individual names).

When you set it up so that the app shows and updates the graphs you will be able to take the next step and automate the system. Why not? It is free anyway!

Steps required for automatic alerting and daily Email reports:
  - use your Google account (create one if necessary), go to https://docs.google.com/spreadsheets and create a sheet called 'SensorsData'
  - go into your SwitchBot app, then into Profile/Preferences/About/Developer Options
  - get a token string (you will not need a Secret Key)
  - in Google Sheet select "Extensions"/"Apps script" and paste there the script code from "hivemonitor.as" file
  - edit the top of the script and insert your token into this line: const token = 'your token here';
  - edit the script to insert your Email (daily reports will be sent there) - it works best with gmail (which expands attached graphs so you can see them instantly)
    Note: _only devices with a "hive" in their name (as given in the app) will be generating alerts about extreme temperature and humidity_
  - edit the script to adjust the critical levels of temperature and humidity
  - save the script (an icon at the top which looks like a small floppy disk; if you know what it looks like :-)
  - press the Triggers icon on the left side (it looks like an alarm clock)
    - create a trigger to run logSensorData() every 30 min (or every hour - this may allow sensors' batteries to run a bit longer)
    - create a trigger to run createDailyCharts() once a day (I set it to just after midnight so that graphs reflect 24h of the last day)
  - to test how the script works you can manually use "Run" icon at the top while selecting either logSensorData() or createDailyCharts() to the right of the "Run" icon
    (if any error occurs it will be reported at the bottom).

If you need help, find any bug, have any comments or ideas, please Email me - you can find the address at the bottom of "hivemonitor.pdf".
