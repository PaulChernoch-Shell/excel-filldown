var XLSX = require('xlsx');

// This test script expects to find a test data file named "test.xlsx" in the current directory.
// The file must have a tab named "Tasks".
// Though not required, assume that the columns are:
//    SubArchetype
//    Phase
//    Activity
//    Method
//    Task
//    Name
//    Duration
//    FixedCost
//    VariableCost

// It loads the spreadsheet, parses the data into JSON, fills in the missing properties and serializes
// the result to an indented JSON string. 

// If rec1 has a property that rec2 is missing, copy the value of the property from rec1 to rec2.
const fillDownRow = (rec1, rec2) => {
    for (const prop in rec1) {
        if (!rec2.hasOwnProperty(prop)) {
            rec2[prop] = rec1[prop];
        }
    }
    return rec2;
};

// Fill-down all missing values from the previous record for all records in the sheet.
//   - Assume all records are in the order they were read from the Excel file.
//   - Asumes that Row 1 of teh sheet held the column headings.
//   - Add an extra property named Row that holds a row number starting at 2.
//   - If multiple records in sequence are all missing a given property, they will all receive
//   - the same value, pulled from the most recent record that has that property.
//   - If the first record is missing any properties, it and possibly subsequent records
//     will end up missing properties.
const fillDownSheet = (records) => {
    if (!records || records.length == 0) { return records; }
    let prevRecord = undefined;
    let rowNumber = 2;
    for (rec of records) {
        rec.Row = rowNumber++;
        if (prevRecord) { 
            fillDownRow(prevRecord, rec);
        }
        prevRecord = rec;
    }
    return records;
};

console.log("Read from an Excel file");
let testExcelFileName = "test.xlsx";
var workbook = XLSX.readFile(testExcelFileName);
var sheet_name_list = workbook.SheetNames;
console.log(`Sheet names: ${sheet_name_list.join(", ")}`);

// Validate that there is a Tab named "Tasks".
let expectedTab = "Tasks";
if (sheet_name_list.indexOf(expectedTab) < 0) {
    console.log(`Did not find a tab named ${expectedTab}`);
    process.exit(1);
}

// Assume that the first row holds column headers. 
// Convert all subsequent rows into records that use the column headers as property names.
let json = XLSX.utils.sheet_to_json(workbook.Sheets[expectedTab]);

// Fill in missing information, assuming that we are loading a hierarchy where repeated
// values for the upper parts of the hierarchy are omitted.
let filledDownJson = fillDownSheet(json);

console.log(`Excel data as JSON: \n${JSON.stringify(filledDownJson, null, 2)}`);

