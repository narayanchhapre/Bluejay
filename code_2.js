const XLSX = require('xlsx');
const moment = require('moment');

// Function to load the Excel file and perform analysis
function analyzeExcelFile(file_path, consecutive_days_threshold = 7) {
    try {
        // Read the Excel file into a DataFrame
        const workbook = XLSX.readFile(file_path);
        const sheet_name_list = workbook.SheetNames;
        const df = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

        // Initialize sets to keep track of printed employees
        const consecutivePrinted = new Set();
        const shortBreakPrinted = new Set();
        const longShiftPrinted = new Set();

        console.log("part 1: ");

        for (let index = 0; index < df.length; index++) {
            const row = df[index];
            const employeeName = row['Employee Name'];
            const positionId = row['Position ID'];

            if (consecutivePrinted.has(employeeName)) {
                continue;
            }

            // Check for consecutive days worked
            if (index > 0 && employeeName === df[index - 1]['Employee Name']) {
                let consecutiveDays = 1;
                for (let i = index - 1; i >= 0; i--) {
                    if (df[i]['Employee Name'] === employeeName) {
                        consecutiveDays += 1;
                    } else {
                        break;
                    }
                }
                if (consecutiveDays >= consecutive_days_threshold) {
                    console.log(`Employee: ${employeeName}, Position: ${positionId}`);
                    consecutivePrinted.add(employeeName);
                }
            }
        }

        console.log("part 2: ");

        const employeeBreaks = {}; // Dictionary to track breaks for each employee

        for (let index = 0; index < df.length; index++) {
            const row = df[index];
            const employeeName = row['Employee Name'];
            const positionId = row['Position ID'];

            if (shortBreakPrinted.has(employeeName)) {
                continue;
            }

            if (employeeName in employeeBreaks) {
                const lastTimeOut = moment(employeeBreaks[employeeName]);
                const timeIn = moment(row['Time']);

                const timeDiff = timeIn.diff(lastTimeOut, 'hours', true);
                if (1 < timeDiff && timeDiff < 10) {
                    console.log(`Employee: ${employeeName}, Position: ${positionId}`);
                    shortBreakPrinted.add(employeeName);
                }
            }

            employeeBreaks[employeeName] = row['Time Out'];
        }

        console.log("part 3: ");

        for (let index = 0; index < df.length; index++) {
            const row = df[index];
            const employeeName = row['Employee Name'];
            const positionId = row['Position ID'];

            if (longShiftPrinted.has(employeeName)) {
                continue;
            }

            // Check for shifts longer than 14 hours
            const durationStr = row['Timecard Hours (as Time)'];
            if (durationStr !== undefined) {
                try {
                    const [hours, minutes] = durationStr.split(':').map(Number);
                    const duration = moment.duration({ hours, minutes });

                    if (duration.asHours() > 14) {
                        console.log(`Employee: ${employeeName}, Position: ${positionId}`);
                        longShiftPrinted.add(employeeName);
                    }
                } catch (error) {
                    // Handle invalid duration format
                    console.error(error);
                }
            }
        }

    } catch (error) {
        console.error(`An error occurred: ${error.message}`);
    }
}

// Example usage
const file_path = 'Assignment_Timecard.xlsx';
analyzeExcelFile(file_path, 7);
