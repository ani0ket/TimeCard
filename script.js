const ExcelJS = require("exceljs");

async function analyzeEmployeeShifts(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);

  let employeeData = {};

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const positionId = row.getCell(1).text;
      const employeeName = row.getCell(8).text;
      const timeIn = row.getCell(3).value;
      const timeOut = row.getCell(4).value;

      if (timeIn && timeOut && positionId && employeeName) {
        if (!employeeData[employeeName]) {
          employeeData[employeeName] = {
            positionId: positionId,
            shifts: [],
          };
        }
        employeeData[employeeName].shifts.push({ timeIn, timeOut });
      }
    }
  });

  let results = {
    "7_consecutive_days": [],
    less_than_10_hours_between_shifts: [],
    more_than_14_hours_shift: [],
  };

  for (let employeeName in employeeData) {
    let shifts = employeeData[employeeName].shifts;
    shifts.sort((a, b) => a.timeIn - b.timeIn);

    let consecutiveDays = 1;
    let previousDay = null;
    let lastTimeOut = null;

    for (let i = 0; i < shifts.length; i++) {
      const shift = shifts[i];
      const shiftDuration = (shift.timeOut - shift.timeIn) / 3600000;

      if (shiftDuration > 14) {
        results.more_than_14_hours_shift.push(
          `${employeeName} (${employeeData[employeeName].positionId})`
        );
      }

      const currentDay = shift.timeIn.toISOString().split("T")[0];
      if (
        previousDay &&
        currentDay ===
          new Date(new Date(previousDay).getTime() + 86400000)
            .toISOString()
            .split("T")[0]
      ) {
        consecutiveDays++;
        if (consecutiveDays >= 7) {
          results["7_consecutive_days"].push(
            `${employeeName} (${employeeData[employeeName].positionId})`
          );
          break;
        }
      } else {
        consecutiveDays = 1;
      }
      previousDay = currentDay;

      if (
        lastTimeOut &&
        shift.timeIn - lastTimeOut < 36000000 &&
        shift.timeIn - lastTimeOut > 3600000
      ) {
        results.less_than_10_hours_between_shifts.push(
          `${employeeName} (${employeeData[employeeName].positionId})`
        );
        break;
      }
      lastTimeOut = shift.timeOut;
    }
  }

  return results;
}

const filePath = "./test.xlsx";

analyzeEmployeeShifts(filePath)
  .then((results) => {
    console.log(
      "Employees who have worked for 7 consecutive days:",
      results["7_consecutive_days"]
    );
    console.log(
      "Employees who have less than 10 hours of time between shifts:",
      results["less_than_10_hours_between_shifts"]
    );
    console.log(
      "Employees who have worked for more than 14 hours in a single shift:",
      results["more_than_14_hours_shift"]
    );
  })
  .catch((error) => {
    console.error("Error:", error);
  });
