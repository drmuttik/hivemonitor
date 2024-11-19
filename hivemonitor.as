//obtain this from "Developer options" in switchbot app
const token = 'your token here';

//only devices with a string 'hive' in the name will generate Email alerts when temperature or humidity is outside of acceptable range
const email = 'your gmail here';

// Define acceptable ranges
const TEMP_RANGE = { min: 10, max: 32 }; //side frame with the sensor is well ouside of the bees' cluster so can be pretty cool
const HUMIDITY_RANGE = { min: 50, max: 75 };

/////////////////////////////////////////////////////////////////////////////////////////
const API_TOKEN = token;

function logSensorData() {
    const url = 'https://api.switch-bot.com/v1.0/devices'; // URL to fetch device list (we use API 1.0; API 2.0 is problematic)
    const headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json"
    };

    const response = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: headers
    });

    const deviceList = JSON.parse(response.getContentText()).body.deviceList;

    // Prepare to log data
    const deviceData = {};
    
    deviceList.forEach(device => {
        if (device.enableCloudService) {
            const deviceId = device.deviceId;
            const deviceName = device.deviceName;

            const statusUrl = 'https://api.switch-bot.com/v1.0/devices/${deviceId}/status';
            const statusResponse = UrlFetchApp.fetch(statusUrl, {
                method: "GET",
                headers: headers
            });
            const statusData = JSON.parse(statusResponse.getContentText()).body;

            // Collect temperature and humidity
            deviceData[deviceName] = {
                temperature: statusData.temperature || 'N/A',
                humidity: statusData.humidity || 'N/A'
            };
        }
    });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SensorsData');
        
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    let rowData = [timestamp];
    let outOfRange = false;  // Flag to check if any data is out of range
    let outOfRangeDetails = []; // Array to store out-of-range details for email
    
    devices.forEach((deviceName, index) => {
        const data = deviceData[deviceName] || { temperature: 'N/A', humidity: 'N/A' };
        const temperature = parseFloat(data.temperature);
        const humidity = parseFloat(data.humidity);
        rowData.push(`${temperature.toFixed(1)}C ${humidity}%`);

        // Only apply out-of-range checks for devices with "hive" in the name
        if (deviceName.toLowerCase().includes("hive")) {
            // Check if temperature or humidity is out of the specified range
            if ((temperature < TEMP_RANGE.min || temperature > TEMP_RANGE.max) || (humidity < HUMIDITY_RANGE.min || humidity > HUMIDITY_RANGE.max)) {
                outOfRange = true; // Set the flag if either value is out of range

                // Highlight the current cell (index + 2 because column A is timestamp)
                const cell = sheet.getRange(sheet.getLastRow() + 1, index + 2); // Next row, device column (B, C, D, etc.)
                cell.setBackground('pink');

                // Collect details of out-of-range events for email
                if (temperature < TEMP_RANGE.min || temperature > TEMP_RANGE.max) {
                    outOfRangeDetails.push(`${deviceName} - Temperature: ${temperature.toFixed(1)}C (out of range ${TEMP_RANGE.min}C-${TEMP_RANGE.max}C)`);
                }
                if (humidity < HUMIDITY_RANGE.min || humidity > HUMIDITY_RANGE.max) {
                    outOfRangeDetails.push(`${deviceName} - Humidity: ${humidity}% (out of range ${HUMIDITY_RANGE.min}%-${HUMIDITY_RANGE.max}%)`);
                }
            }
        }
    });

    // Append the row data to the sheet
    sheet.appendRow(rowData);

    // Get the index of the last row (which is the row we just appended)
    var lastRow = sheet.getLastRow();

    // Specify the range to be right-justified (columns B through G in this example)
    var range = sheet.getRange(lastRow, 2, 1, rowData.length - 1); // Assumes first column is the timestamp

    // Set the horizontal alignment to 'right' for the range
    range.setHorizontalAlignment("right");

    // Send an email if any data is out of range for devices with "hive" in the name
    if (outOfRange) {
        const subject = "Alert: Hive sensor data out of range";
        const body = `Some hive sensor readings are out of the acceptable range:\n\n` + outOfRangeDetails.join("\n");
        MailApp.sendEmail(email, subject, body);
    }
}

function createDailyCharts() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SensorsData');
    const dataRange = sheet.getDataRange();
    const dataValues = dataRange.getValues();

    const last24HoursData = [];
    const now = new Date();
    const twentyFourHoursAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // 24 hours ago
    const reportDate = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // Use date -1 for charts and email

    const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

    // Clear previous temp and humidity data columns
    sheet.getRange('Y:AG').clear(); // Clear columns Y to AG for fresh data

    // Extract all rows within the last 24 hours
    for (let i = 1; i < dataValues.length; i++) {  // Skip header
        const rowTimestamp = new Date(dataValues[i][0]); // First column is the timestamp
        if (rowTimestamp >= twentyFourHoursAgo && rowTimestamp <= now) {
            last24HoursData.push(dataValues[i]); // Add the whole row if it's within the last 24 hours
        }
    }

    if (last24HoursData.length === 0) {
        Logger.log("No data found for the last 24 hours.");
        return;
    }

    // Prepare data for temperature and humidity columns Y to AG
    const tempData = [];
    const humidityData = [];

    last24HoursData.forEach(row => {
        const time = new Date(row[0]); // Timestamp

        // Extract temperature and humidity from text
        const [tempLeftHive, humidityLeftHive] = extractTempHumidity(row[1]); // Column B
        const [tempRightHive, humidityRightHive] = extractTempHumidity(row[2]); // Column C
        const [tempNucHive, humidityNucHive] = extractTempHumidity(row[3]); // Column D
        const [tempGarage, humidityGarage] = extractTempHumidity(row[4]); // Column E

        // Push valid temperature data to tempData array
        tempData.push([time, tempLeftHive, tempRightHive, tempNucHive, tempGarage]);
        
        // Push valid humidity data to humidityData array
        humidityData.push([time, humidityLeftHive, humidityRightHive, humidityNucHive]);
    });

    // Write temperature data to columns Y to AD
    if (tempData.length > 0) {
        sheet.getRange(2, 26, tempData.length, tempData[0].length).setValues(tempData); // Y is column 26
    }

    // Write humidity data to columns AE to AG
    if (humidityData.length > 0) {
        sheet.getRange(2, 31, humidityData.length, humidityData[0].length).setValues(humidityData); // AE is column 31
    }

    // Format columns Z and AE as Time and HH:mm
    sheet.getRange(2, 26, tempData.length, 1).setNumberFormat("HH:mm"); // Column Z
    sheet.getRange(2, 31, humidityData.length, 1).setNumberFormat("HH:mm"); // Column AE

    // Clear previous charts
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));

    // Create temperature chart
    const tempChart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sheet.getRange(2, 26, tempData.length, 5)) // Time + 4 temperature columns (Y to AD)
        .setPosition(2, 9, 0, 0)  // Positioning the chart to the right of column I (column 9)
        .setOption('title', `Temperature for ${formattedDate}`)
        .setOption('legend', { position: 'right' })
        .setOption('hAxis', {
            title: 'Time',
            format: 'HH:mm',
            minorGridlines: { count: 1 },
            gridlines: { count: 24 },
            slantedText: true,       // Enable slanted text
            slantedTextAngle: 90     // Set the angle to 90 degrees
            //ticks: generateTicks(last24HoursData[0][0], last24HoursData[last24HoursData.length - 1][0], 1) // Set 1-hour ticks
        })
        .setOption('vAxis', {
            title: 'Temperature (°C)',
            viewWindow: {
                max: Math.ceil(Math.max(...tempData.flatMap(row => row.slice(1))) / 10) * 10,
                min: Math.floor(Math.min(...tempData.flatMap(row => row.slice(1))) / 10) * 10
            }
        })
        .setOption('series', {
            0: { labelInLegend: 'Left hive', lineWidth: 2 },
            1: { labelInLegend: 'Right hive', lineWidth: 2 },
            2: { labelInLegend: 'NUC hive', lineWidth: 2 },
            3: { labelInLegend: 'Garage', lineWidth: 2 }
        })
        .setOption('curveType', 'none')  // Ensure it's a line chart with no smoothing
        .setOption('pointSize', 5) // Add points to make it easier to see
        .build();

    // Create humidity chart
    const humidityChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(2, 31, humidityData.length, 4)) // Time + 3 humidity columns (AE to AG)
    .setPosition(20, 9, 0, 0)  // Positioning the chart below the temperature chart to the right of column I
    .setOption('title', `Humidity for ${formattedDate}`)
    .setOption('legend', { position: 'right' })
      .setOption('hAxis', {
        title: 'Time',
        format: 'HH:mm',
        minorGridlines: { count: 1 },
        gridlines: { count: 24 },
        slantedText: true,       // Enable slanted text
        slantedTextAngle: 90     // Set the angle to 90 degrees
        //ticks: generateTicks(last24HoursData[0][0], last24HoursData[last24HoursData.length - 1][0], 1) // Set 1-hour ticks
      })
    .setOption('vAxis', {
        title: 'Humidity (%)',
        viewWindow: {
            max: Math.ceil(Math.max(...humidityData.flatMap(row => row.slice(1))) / 10) * 10,
            min: Math.floor(Math.min(...humidityData.flatMap(row => row.slice(1))) / 10) * 10
        }
    })
    .setOption('series', {
        0: { labelInLegend: 'Left hive', lineWidth: 2 },
        1: { labelInLegend: 'Right hive', lineWidth: 2 },
        2: { labelInLegend: 'NUC hive', lineWidth: 2 }
    })
    .setOption('curveType', 'none')
    .setOption('pointSize', 5)
    .build();

    // Insert the charts into the sheet
    sheet.insertChart(tempChart);
    sheet.insertChart(humidityChart);

    minChart = plotDailyTemperatures("min");  // Plot daily minimums
    maxChart = plotDailyTemperatures("max");  // Plot daily maximums

    // Prepare the email
    const tempChartBlob = tempChart.getAs('image/png');
    const humidityChartBlob = humidityChart.getAs('image/png');
    const maxChartBlob = maxChart.getAs('image/png');
    const minChartBlob = minChart.getAs('image/png');
    
    const subject = `Temperature and humidity (${formattedDate})`;
    const body = `Temperature and humidity charts for ${formattedDate} and max/min temperatures for the last 31 days.`;
    
    MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        attachments: [
            tempChartBlob.setName(`Temperature_${formattedDate}.png`),
            humidityChartBlob.setName(`Humidity_${formattedDate}.png`),
            maxChartBlob.setName(`MaximumTemperatures.png`),
            minChartBlob.setName(`MinimumTemperatures.png`)
        ]
    });
}

// Helper function to extract temperature and humidity from text
function extractTempHumidity(text) {
    const tempMatch = text.match(/(-?\d+(\.\d+)?)(?=C)/);
    const humidityMatch = text.match(/(\d+)(?=%)/);

    const temperature = tempMatch ? parseFloat(tempMatch[0]) : NaN;
    const humidity = humidityMatch ? parseFloat(humidityMatch[0]) : NaN;

    return [temperature, humidity];
}

function plotDailyTemperatures(statType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SensorsData");
  const data = sheet.getDataRange().getValues();
  
  // Extract date and temperature data (columns 1-5)
  const tempData = data.slice(1).map(row => [row[0], row[1], row[2], row[3], row[4]]);
  
  const tempByDate = {};

  tempData.forEach(row => {
    const date = new Date(row[0]).toDateString();
    if (!tempByDate[date]) tempByDate[date] = [];
    
    row.slice(1).forEach(entry => {
      const temp = parseFloat(entry.split("C")[0]);
      tempByDate[date].push(temp);
    });
  });
  
  // Calculate daily minimum or maximum temperatures based on statType
  const dailyTempData = Object.keys(tempByDate).map(dateStr => {
    const date = new Date(dateStr);
    const dailyTemps = [date]; // Store Date object instead of formatted string
    for (let i = 0; i < 4; i++) {
      const colTemps = tempByDate[dateStr].filter((_, index) => index % 4 === i);
      
      // Correct logic here:
      // Use Math.min for minimum temperatures and Math.max for maximum temperatures
      dailyTemps.push(statType === "min" ? Math.min(...colTemps) : Math.max(...colTemps));
    }
    return dailyTemps;
  });

  // Only keep the last 31 entries for a monthly view
  const last31Entries = dailyTempData.slice(-31);
  
  // Extract dates and temperatures for the last 31 entries
  const dates = last31Entries.map(row => [row[0]]); // Date objects
  const temperatures = last31Entries.map(row => row.slice(1));
  
  // Define destination columns based on statType
  const dateCol = statType === "min" ? 20 : 20;  // Dates in column T
  const tempStartCol = statType === "min" ? 21 : 16; // Min in U-Y, Max in P-S
  
  // Save dates and temperatures to appropriate columns
  const dateRange = sheet.getRange(2, dateCol, dates.length, 1);
  dateRange.setValues(dates);
  dateRange.setNumberFormat("dd-MMM-yy"); // Set desired date format
  
  const tempRange = sheet.getRange(2, tempStartCol, temperatures.length, 4);
  tempRange.setValues(temperatures); // Temperatures in columns P-S or U-Y
  
  // Configure chart settings
  const title = statType === "min" ? "Daily Minimum Temperatures" : "Daily Maximum Temperatures";
  const position = statType === "min" ? [20, 15, 0, 0] : [2, 15, 0, 0]; // Position for min or max chart
  
  const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(2, dateCol, dates.length, 1))       // X-axis from column T
      .addRange(sheet.getRange(2, tempStartCol, temperatures.length, 4)) // Y-axis from columns P-S or U-Y
      .setPosition(...position)
      .setOption('title', title)
      .setOption('hAxis', {
          title: 'Date',
          slantedText: true,
          slantedTextAngle: 90,
          format: 'dd-MMM-yy', // Correct format for display
          gridlines: {
              count: last31Entries.length   // One gridline per day
          }
      })
      .setOption('vAxis', {
          title: 'Temperature (°C)',
          gridlines: { count: -1 },
          minorGridlines: { count: 1 }
      })
      .setOption('legend', { position: 'none' })
      .setOption('series', {
          0: { pointSize: 5 },
          1: { pointSize: 5 },
          2: { pointSize: 5 },
          3: { pointSize: 5 }
      })
      .build();
  
  sheet.insertChart(chart);
  return(chart);
}
