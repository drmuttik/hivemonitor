//obtain this from "Developer options" in switchbot app
const token = 'your token here';

//only devices with a string 'hive' in the name will generate Email alerts when temperature or humidity is outside of acceptable range
const email = 'your gmail here';

// Define acceptable ranges
const TEMP_RANGE = { min: 10, max: 32 }; //side frame with the sensor is well ouside of the bees' cluster so can be pretty cool
const HUMIDITY_RANGE = { min: 50, max: 75 };

// Define one sensor measuring temperature outside (name)
const OUTSIDE_SENSOR="Porch";

// Set to true to include the outside humidity in the humidity chart (to exclude - set to false)
const PLOT_OUTSIDE_HUMIDITY = true;

// Maximum number of email alerts per day
const MAX_EMAIL_ALERTS_PER_DAY = 5;

///////////////////////////////////////////////////////////////////////////////////////

// Global variable to track email count and date
let emailCount = 0;
let lastEmailResetDate = '';

function logSensorData() {
  const url = 'https://api.switch-bot.com/v1.0/devices';
  const headers = {
    "Authorization": token,
    "Content-Type": "application/json"
  };
  
  let response;
  for (let attempt = 1; attempt <= 10; attempt++) { //10 retries
    try {
      response = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: headers
      });
      break; // Exit the loop if successful
    } catch (error) {
      if (attempt < 10 ) {
        Utilities.sleep(60000); // Wait time in milliseconds
      } else {
        console.error("Max retries reached. Exiting.");
      }
    }
  }

    const deviceList = JSON.parse(response.getContentText()).body.deviceList;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // Reset email count at midnight
    const todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (lastEmailResetDate !== todayDate) {
        emailCount = 0;
        lastEmailResetDate = todayDate;
    }

    // Prepare header row, excluding "Hub Mini" devices
    const headerRow = ["Timestamp"];
    const deviceIdMap = {}; // Map to track device IDs for correct column placement

    const filteredDeviceList = deviceList.filter(device => {
        const deviceName = device.deviceName || '';
        if (deviceName.trim() !== '' && !deviceName.toLowerCase().includes("hub mini")) {
            headerRow.push(device.deviceName);
            deviceIdMap[device.deviceId] = headerRow.length - 1; // Map device ID to column index
            return true;
        }
        return false;
    });

    //check if there is at least one device with "hive" in the name (commented out to allow people not assigning any names)
    //const hasHive = filteredDeviceList.some(device => device.deviceName.toLowerCase().includes("hive"));
    //if (!hasHive) {
    //  throw new Error('No hives found.');
    //}

    // Check if column A is unpainted ("none") - this ensures the following block is executed once (like an initialization function)
    const columnAFirstCell = sheet.getRange(1, 1);
    const columnAFirstCellBackground = columnAFirstCell.getBackground();
    if (columnAFirstCellBackground === "none" || columnAFirstCellBackground === "#ffffff") {
      // Paint first column light green
      const columnARange = sheet.getRange(1, 1, sheet.getMaxRows(), 1);
      columnARange.setBackground('lightgreen');

      // Paint even columns (2, 4, ...) lighter grey, only if they contain data
      for (let col = 2; col <= headerRow.length; col += 2) {
        const columnRange = sheet.getRange(1, col, sheet.getMaxRows(), 1);
        const columnValues = columnRange.getValues();

        // Check if there is data in the column
        const hasData = columnValues.some(row => row[0] !== "" && row[0] !== null);

        if (hasData) {
          columnRange.setBackground('#E8E8E8'); // Lighter grey
        }
      }

      // Paint header row yellow and assign comments
      const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
      headerRange.setValues([headerRow])
          .setFontWeight('bold')
          .setFontSize(headerRange.getFontSize() + 1)
          .setBackground('yellow');

      // Add device IDs as comments in the header row
      Object.keys(deviceIdMap).forEach(deviceId => {
        const columnIndex = deviceIdMap[deviceId] + 1;
        sheet.getRange(1, columnIndex).setComment(`Device ID: ${deviceId}`);
      });
    }

    // Prepare to log data
    const deviceData = {};
    filteredDeviceList.forEach(device => {
        if (device.enableCloudService) {
            const deviceId = device.deviceId;
            const statusUrl = `https://api.switch-bot.com/v1.0/devices/${deviceId}/status`;
            const statusResponse = UrlFetchApp.fetch(statusUrl, {
                method: "GET",
                headers: headers
            });
            const statusData = JSON.parse(statusResponse.getContentText()).body;

            // Collect temperature and humidity
            deviceData[deviceId] = {
                temperature: statusData.temperature || 'N/A', 
                humidity: statusData.humidity || 'N/A'
            };
        }
    });

    // Add timestamp and device data to the sheet
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    let rowData = Array(headerRow.length).fill(''); // Initialize row with empty values
    rowData[0] = timestamp; // Set timestamp in column A

    let outOfRangeDetails = []; // Store details for out-of-range alerts

    Object.keys(deviceData).forEach(deviceId => {
        const columnIndex = deviceIdMap[deviceId];
        if (columnIndex !== undefined) {
            const data = deviceData[deviceId];
            const temperature = parseFloat(data.temperature);
            const humidity = parseFloat(data.humidity);
            // Replace NaN with default values
            if (isNaN(temperature)) temperature = 0;
            if (isNaN(humidity)) humidity = 0;
            rowData[columnIndex] = `${temperature.toFixed(1)}C ${humidity}%`;

            // Check for out-of-range values only for devices with "hive" in the name
            const deviceName = headerRow[columnIndex];
            if (deviceName.toLowerCase().includes("hive")) {
                if ((temperature < TEMP_RANGE.min || temperature > TEMP_RANGE.max) || 
                    (humidity < HUMIDITY_RANGE.min || humidity > HUMIDITY_RANGE.max)) {
                    // Highlight the cell
                    const cell = sheet.getRange(sheet.getLastRow() + 1, columnIndex + 1); // Next row, respective column
                    cell.setBackground('pink');

                    // Add details to the out-of-range list
                    if (temperature < TEMP_RANGE.min || temperature > TEMP_RANGE.max) {
                        outOfRangeDetails.push(`${deviceName} - Temperature: ${temperature.toFixed(1)}C (Range: ${TEMP_RANGE.min}-${TEMP_RANGE.max}C)`);
                    }
                    if (humidity < HUMIDITY_RANGE.min || humidity > HUMIDITY_RANGE.max) {
                        outOfRangeDetails.push(`${deviceName} - Humidity: ${humidity}% (Range: ${HUMIDITY_RANGE.min}-${HUMIDITY_RANGE.max}%)`);
                    }
                }
            }
        }
    });

    // Append the data to the sheet
    sheet.appendRow(rowData);

    // Send email alert if any out-of-range values are found
    if (outOfRangeDetails.length > 0 && emailCount < MAX_EMAIL_ALERTS_PER_DAY) {
        const subject = 'HiveMonitor Alert: Out-of-Range Values Detected';
        const body = `Timestamp: ${timestamp}\n\nThe following values are out of range:\n\n${outOfRangeDetails.join('\n')}`;
        GmailApp.sendEmail(email, subject, body);
        emailCount++;
    }
}

function createDailyCharts() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const dataRange = sheet.getDataRange();
    const dataValues = dataRange.getValues();
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const now = new Date();
    const twentyFourHoursAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000));
    const reportDate = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // Use date -1 for charts
    const formattedDate = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

    // Clear previous charts and data
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));

    // Step 1: Count columns with data and wipe old columns starting from NDATA+1+8
    const NDATA = headerRow.filter(cell => cell !== "").length; // Count non-empty columns
    
    // Step 2: Use column NDATA+8 for graph timestamps
    const timestampCol = NDATA + 8;
    sheet.getRange(100, timestampCol, 64, NDATA + 1).clear(); // Wipe old data

    const last24HoursData = [];
    for (let i = 1; i < dataValues.length; i++) { // Skip header
        const rowTimestamp = new Date(dataValues[i][0]);
        if (rowTimestamp >= twentyFourHoursAgo && rowTimestamp <= now) {
            last24HoursData.push(dataValues[i]);
        }
    }
    if (last24HoursData.length === 0) {
        Logger.log("No data found for the last 24 hours.");
        return;
    }

    const tempData = [];
    const humidityData = [];
    const hiveColumns = [];
    let outsideSensorColumn = -1;

    // Step 3: Count columns with "hive" in the name
    headerRow.forEach((name, colIndex) => {
        const lowerName = name.toLowerCase();
        if (lowerName.includes("hive")) {
            hiveColumns.push(colIndex);
        } else if (name === OUTSIDE_SENSOR) {
            outsideSensorColumn = colIndex;
        }
    });

    // Step 4: Process data and populate temperature and humidity data
    last24HoursData.forEach(row => {
        const time = row[0];
        const tempRow = [time];
        const humidityRow = [time];

        hiveColumns.forEach(colIndex => {
            const [temp, humidity] = extractTempHumidity(row[colIndex]);
            tempRow.push(temp);
            humidityRow.push(humidity);
        });

        if (outsideSensorColumn !== -1) {
            const [tempOutside, humidityOutside] = extractTempHumidity(row[outsideSensorColumn]);
            tempRow.push(tempOutside);
            if (PLOT_OUTSIDE_HUMIDITY) {
                humidityRow.push(humidityOutside); // Add humidity data for external sensor if PLOT_OUTSIDE_HUMIDITY is true
            }
        }

        tempData.push(tempRow);
        humidityData.push(humidityRow);
    });

    // Step 5: Write temperature data to columns starting at NDATA+1+9
    const tempStartCol = NDATA + 9;
    const tempRange = sheet.getRange(2, tempStartCol, tempData.length, tempData[0].length);
    tempRange.setValues(tempData);

    // Convert timestamps to HH:mm format in the temp data range
    const timestampRange = sheet.getRange(2, tempStartCol, tempData.length, 1); // Get the timestamp column
    const timestampValues = timestampRange.getValues();
    for (let i = 0; i < timestampValues.length; i++) {
        const formattedTimestamp = Utilities.formatDate(timestampValues[i][0], Session.getScriptTimeZone(), "HH:mm");
        timestampValues[i][0] = formattedTimestamp;
    }
    timestampRange.setValues(timestampValues); // Update the column with formatted timestamps

    // Step 6: Format temperature data columns as numbers
    sheet.getRange(2, tempStartCol + 1, tempData.length, tempData[0].length - 1).setNumberFormat("0.0");

    // Create temperature chart
    const tempChart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(tempRange)
        .setPosition(2, NDATA+1, 0, 0) // Set position to (2, 9)
        .setOption('title', `Temperature for ${formattedDate}`)
        .setOption('legend', { position: 'right' })
        .setOption('hAxis', {
            title: 'Time',
            format: 'HH:mm', // Ensure that the time format is "HH:mm"
            gridlines: { count: 6 },
            ticks: timestampValues.map(row => row[0]), // Use formatted timestamps for ticks
        })
        .setOption('vAxis', {
            title: 'Temperature (°C)',
            viewWindow: {
                max: Math.ceil(Math.max(...tempData.flatMap(row => row.slice(1))) / 10) * 10,
                min: Math.floor(Math.min(...tempData.flatMap(row => row.slice(1))) / 10) * 10
            }
        })
        .setOption('pointSize', 5) // Add points to the lines
        .setOption('series', hiveColumns.concat(outsideSensorColumn !== -1 ? [outsideSensorColumn] : []).reduce((series, colIndex, i) => {
            series[i] = { labelInLegend: headerRow[colIndex] };
            return series;
        }, {}))
        .build();

    sheet.insertChart(tempChart);

    // Step 7: Wipe and populate columns for Humidity graph
    const humidityStartCol = tempStartCol + tempData[0].length; // After temp data columns
    const humidityRange = sheet.getRange(2, humidityStartCol, humidityData.length, humidityData[0].length);
    humidityRange.setValues(humidityData);

    // Convert timestamps to HH:mm format for humidity graph
    const humidityTimestampRange = sheet.getRange(2, humidityStartCol, humidityData.length, 1); // Get the timestamp column
    const humidityTimestampValues = humidityTimestampRange.getValues();
    for (let i = 0; i < humidityTimestampValues.length; i++) {
        const formattedTimestamp = Utilities.formatDate(humidityTimestampValues[i][0], Session.getScriptTimeZone(), "HH:mm");
        humidityTimestampValues[i][0] = formattedTimestamp;
    }
    humidityTimestampRange.setValues(humidityTimestampValues); // Update the column with formatted timestamps

    // Step 8: Format humidity data columns as numbers
    sheet.getRange(2, humidityStartCol + 1, humidityData.length, humidityData[0].length - 1).setNumberFormat("0.0");

    // Create humidity chart
    const humidityChart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(humidityRange)
        .setPosition(20, NDATA+1, 0, 0) // Set position to (20, 9)
        .setOption('title', `Humidity for ${formattedDate}`)
        .setOption('legend', { position: 'right' })
        .setOption('hAxis', {
            title: 'Time',
            format: 'HH:mm', // Ensure that the time format is "HH:mm"
            gridlines: { count: 6 },
            ticks: humidityTimestampValues.map(row => row[0]), // Use formatted timestamps for ticks
        })
        .setOption('vAxis', {
            title: 'Humidity (%)',
            viewWindow: {
                max: Math.ceil(Math.max(...humidityData.flatMap(row => row.slice(1))) / 10) * 10,
                min: Math.floor(Math.min(...humidityData.flatMap(row => row.slice(1))) / 10) * 10
            }
        })
        .setOption('pointSize', 5) // Add points to the lines
        .setOption('series', hiveColumns.concat(PLOT_OUTSIDE_HUMIDITY ? [outsideSensorColumn] : []).reduce((series, colIndex, i) => {
            series[i] = { labelInLegend: headerRow[colIndex] };
            return series;
        }, {}))
        .build();

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

    const temperature = tempMatch ? parseFloat(tempMatch[0]) : 0; //to process records like "NaNC 80%"
    const humidity = humidityMatch ? parseFloat(humidityMatch[0]) : 0;

    return [temperature, humidity];
}

function plotDailyTemperatures(statType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const headerRow = data[0]; // Header row (first row)
  
  // Extract relevant columns
  const timestampCol = 0; // Timestamp is in the first column
  const hiveColumns = [];
  let outsideSensorColumn = -1;

  headerRow.forEach((header, index) => {
    if (header.toLowerCase().includes("hive")) {
      hiveColumns.push(index);
    } else if (header.includes(OUTSIDE_SENSOR)) {
      outsideSensorColumn = index;
    }
  });

  //console.log(`Stat type: ${statType}`);
  //console.log(`Hive columns: ${hiveColumns}`);
  //console.log(`Outside sensor column: ${outsideSensorColumn}`);

  // Prepare data
  const tempData = data.slice(1).map(row => {
    const date = new Date(row[timestampCol]);
    if (isNaN(date)) return null;

    const tempRow = [date];
    hiveColumns.forEach(colIndex => {
      const temp = parseFloat(row[colIndex]?.split("C")[0]);
      tempRow.push(temp);
    });

    if (outsideSensorColumn !== -1) {
      const tempOutside = parseFloat(row[outsideSensorColumn]?.split("C")[0]);
      tempRow.push(tempOutside);
    }

    return tempRow;
  }).filter(row => row !== null);

  // Group by date
  const tempByDate = {};
  tempData.forEach(row => {
    const dateStr = row[0].toDateString();
    if (!tempByDate[dateStr]) tempByDate[dateStr] = [];
    row.slice(1).forEach(temp => tempByDate[dateStr].push(temp));
  });

  const dailyTempData = Object.keys(tempByDate).map(dateStr => {
    const date = new Date(dateStr);
    const dailyTemps = [date];

    // Group temperatures by columns
    const columnsData = Array(hiveColumns.length + (outsideSensorColumn !== -1 ? 1 : 0)).fill().map(() => []);

    tempByDate[dateStr].forEach((temp, index) => {
      // Convert invalid values to 0
      const numericValue = isNaN(temp) || temp === null ? 0 : temp;
      columnsData[index % columnsData.length].push(numericValue);
    });

    // Calculate statType for each column
    columnsData.forEach(colTemps => {
      if (colTemps.length > 0) {
        dailyTemps.push(statType === "min" ? Math.min(...colTemps) : Math.max(...colTemps));
      } else {
        dailyTemps.push(0); // No data for this column, use 0 as a default
      }
    });

    //console.log(`Date: ${dateStr}, ${statType} temperatures: ${dailyTemps.slice(1)}`);
    return dailyTemps;
  });

  // Format for the last 31 days
  const last31Entries = dailyTempData.slice(-31);
  const dates = last31Entries.map(row => [row[0]]);
  const temperatures = last31Entries.map(row => row.slice(1));

  if (dates.length === 0 || temperatures.length === 0) {
    console.log("No data to write or plot.");
    return;
  }

  // Write to sheet
  const NDATA = headerRow.filter(cell => cell !== "").length; // Count non-empty columns
  const rowStart = statType === "min" ? 100 : 132; // Minimums at row 100, maximums at row 132
  const colStart = NDATA+9; // Start with a column safely after sensors data
  const dateCol = colStart; // Column for dates
  const tempStartCol = colStart + 1; // Column for temperatures

  const dateRange = sheet.getRange(rowStart, dateCol, dates.length, 1);
  dateRange.setValues(dates);
  dateRange.setNumberFormat("dd-MMM-yy");

  const tempRange = sheet.getRange(rowStart, tempStartCol, temperatures.length, temperatures[0].length);
  tempRange.setValues(temperatures);
  tempRange.setNumberFormat("0.0");

  // Create chart
  
  const title = statType === "min" ? "Daily Minimum Temperatures" : "Daily Maximum Temperatures";
  const chartPosition = statType === "min" ? [20, NDATA+6, 0, 0] : [2, NDATA+6, 0, 0];

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(rowStart, dateCol, dates.length, 1)) // X-axis
    .addRange(sheet.getRange(rowStart, tempStartCol, temperatures.length, temperatures[0].length)) // Y-axis
    .setPosition(...chartPosition)
    .setOption('title', title)
    .setOption('hAxis', {
      title: 'Date',
      slantedText: true,
      slantedTextAngle: 90,
      format: 'dd-MMM-yy',
      gridlines: { count: last31Entries.length }
    })
    .setOption('vAxis', {
      title: 'Temperature (°C)',
      gridlines: { count: -1 },
      minorGridlines: { count: 1 }
    })
    .setOption('legend', { position: 'none' })
    .setOption('pointSize', 5)
    .build();

  sheet.insertChart(chart);
  return(chart);
}
