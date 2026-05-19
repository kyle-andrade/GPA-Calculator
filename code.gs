/**
 * Author: Kyle Andrade
 * GPA Calculator
 * 
 * Values are inserted in Class Records Sheet.
 * The Grade Cell should have a valid grade variable from the GradeVariables sheet OR PE (Planned Enrollment) OR FE (Future Enrollment)
 * Data will be transfered automatically tot he Simplified Records which is a smaller data set from the Class Records.
 * GPA Data contains the Total GPA and the GPA for each School.
 * GradeVariables contains the grade point values from each school.
 */

function calculateGPA() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Class Records");
  var simplifiedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Simplified Records");
  var data = sheet.getDataRange().getValues();
  var gradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GradeVariables");
  var gpaBySemster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GPA By Semester")

  var schoolOneName = gradeSheet.getRange("A1").getValues(); // Name of Grade One School
  var schoolTwoName = gradeSheet.getRange("D1").getValues(); // Name of Grade Two School

  var schoolOneGradeData = gradeSheet.getRange("A2:B6").getValues(); // Grade scale lookup table for School 1
  var schoolTwoGradeData = gradeSheet.getRange("D2:E11").getValues(); // Grade scale lookup table for School 2
  // var gradeData = gradeSheet.getRange("A1:B5").getValues(); // Grade scale lookup table

  var GPADataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GPA Data"); // GPA Data Sheet
  
  var semesterGPA = [];
  var semesterPosition = 0;
  var semesterClassAmount = 0;

  var totalCreditHours = 0;
  var totalWeightedPoints = 0;
  var totalEarnedCredits = 0;
  var totalCreditsInProgress = 0;
  var totalFutureEnrolled = 0;
  var totalFuturePlanned = 0;
  
  var schoolOneGPA = 0;
  var schoolOneRoundedGPA = 0;
  var schoolOneCreditHours = 0;
  var schoolOneWeightedPoints = 0;
  var schoolOneCreditsInProgress = 0;
  var schoolOneEarnedCredits = 0;
  var schoolOneFutureEnrolled = 0;
  var schoolOnePlannedEnrolled = 0;

  var schoolTwoGPA = 0;
  var schoolTwoRoundedGPA = 0;
  var schoolTwoCreditHours = 0;
  var schoolTwoWeightedPoints = 0;
  var schoolTwoCreditsInProgress = 0;
  var schoolTwoEarnedCredits = 0;
  var schoolTwoFutureEnrolled = 0;
  var schoolTwoPlannedEnrolled = 0;

  for (var i = 1; i < data.length; i++) { // Skip headers
    var creditHours = parseFloat(data[i][7]); // Column H (Credit Hours)
    var gradeReceived = data[i][8]; // Column I (Grade Received)
    var school = data[i][1]; // Column B (School)

    var gradePoints;

    /*
    // Start Comment \\
    // Not needed for now.
    // Dropdown is inputed manually
    var schoolDropdownValues = [
    schoolOneName,
    schoolTwoName
    ]
    var schoolRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(schoolDropdownValues, true)
        .build();

    //sheet.getRange("B2:B").setDataValidation(schoolRule);
    // End Comment \\
*/

    if (i === 1 || sheet.getRange(i, 1).getValue() !== sheet.getRange(i - 1, 1).getValue() || sheet.getRange(i - 1, 1).getValue() === "Semester") {
      // New Semester
      if (semesterClassAmount > 0) {
        semesterGPA[semesterPosition] = semesterGPA[semesterPosition] / semesterClassAmount;
      }
      semesterPosition += 1;
      semesterGPA[semesterPosition] = sheet.getRange(i, 10).getValue() || 0;
      semesterClassAmount = 1;
    } else {
      // Same Semester
      semesterGPA[semesterPosition] += sheet.getRange(i, 10).getValue() || 0;
      semesterClassAmount += 1;
    }
    
    //console.log("Semester GPA", semesterGPA);

    if (school == schoolOneName) {
      //sheet.getRange([i][8]).setDataValidation(schoolOneGradeData);
      
      gradePoints = lookupGradePoints(gradeReceived, schoolOneGradeData); // Get Grade Points for school 1
    } else if (school == schoolTwoName) {
      //sheet.getRange([i][8]).setDataValidation(schoolTwoGradeData);
      
      gradePoints = lookupGradePoints(gradeReceived, schoolTwoGradeData); // Get Grade Points for school 2
    } // End else/if
    
    var weightedPoints = gradePoints * creditHours; // Calculate Weighted Points

    if (!isNaN(creditHours)) {
      totalCreditHours += creditHours;

      if (school == schoolOneName) {
        schoolOneCreditHours += creditHours;
      } else if (school == schoolTwoName) {
        schoolTwoCreditHours += creditHours;
      } // End if else [finding school]

      if (gradeReceived !== "" && gradeReceived !== "FE" && gradeReceived !== "PE") { 
        totalEarnedCredits += creditHours;
        totalWeightedPoints += weightedPoints;
        if (school == schoolOneName) {
          schoolOneEarnedCredits += creditHours;
          schoolOneWeightedPoints += weightedPoints;
        } else if (school == schoolTwoName) {
          schoolTwoEarnedCredits += creditHours;
          schoolTwoWeightedPoints += weightedPoints;
        } // End if else [finding school]
      } else if (gradeReceived === "FE") {
        totalFutureEnrolled += creditHours;
        if (school == schoolOneName) {
          schoolOneFutureEnrolled += creditHours;
        } else if (school == schoolTwoName) {
          schoolTwoFutureEnrolled += creditHours;
        } // End if else [finding school]
      } else if (gradeReceived === "PE") {
        totalFuturePlanned += creditHours;
        if (school == schoolOneName) {
          schoolOnePlannedEnrolled += creditHours;
        } else if (school == schoolTwoName) {
          schoolTwoPlannedEnrolled += creditHours;
        } // End if else [finding school]
      } else {
        totalCreditsInProgress += creditHours;
        if (school == schoolOneName) {
          schoolOneCreditsInProgress += creditHours;
        } else if (school == schoolTwoName) {
          schoolTwoCreditsInProgress += creditHours;
        } // End if else [finding school]
      } // End if - else if - else (gradeRecieved)
    } // End if (!isNaN(creditHours))

    // Update Grade Points and Weighted Points in Sheet
    sheet.getRange(i + 1, 10).setValue(gradePoints); // Column J (Grade Points)
    sheet.getRange(i + 1, 11).setValue(weightedPoints); // Column K (Weighted Points)
  } // End for (i < data.length)

  var gpaROW = 2;
  for (var i = 2; i < semesterGPA.length; i++) {
    if (semesterGPA[i] === "" || semesterGPA[i] === " " || semesterGPA[i] === NaN) {
      continue;
    } else {
      gpaBySemster.getRange(gpaROW, 2).setValue(semesterGPA[i]);
      gpaROW += 1;
    }
  }

  var totalGPA = totalWeightedPoints / totalEarnedCredits || 0;
  var roundedGPA = Math.round(totalGPA * 100) / 100;

  schoolOneGPA = schoolOneWeightedPoints / schoolOneEarnedCredits || 0;
  schoolOneRoundedGPA = Math.round(schoolOneGPA * 100) / 100;

  schoolTwoGPA = schoolTwoWeightedPoints / schoolTwoEarnedCredits || 0;
  schoolTwoRoundedGPA = Math.round(schoolTwoGPA * 100) / 100;

  // Update Totals
  sheet.getRange("M2").setValue(totalCreditHours); // Total Credit Hours
  sheet.getRange("N2").setValue(totalWeightedPoints); // Total Weighted Points
  sheet.getRange("M5").setValue(totalGPA); // Total GPA
  sheet.getRange("N5").setValue(roundedGPA); // Rounded GPA
  sheet.getRange("M8").setValue(totalEarnedCredits); // Total Earned Credits
  sheet.getRange("N8").setValue(totalCreditsInProgress); // Total Credits In-Progress
  sheet.getRange("M11").setValue(totalFutureEnrolled); // Total Credits that are Enrolled
  sheet.getRange("N11").setValue(totalFuturePlanned); // Total Credits that are Planned but NOT Enrolled for or Future Enrolled

  GPADataSheet.getRange("B2").setValue(totalGPA); // Total GPA
  GPADataSheet.getRange("B3").setValue(schoolOneGPA); // School 1 GPA
  GPADataSheet.getRange("B4").setValue(schoolTwoGPA); // School 2 GPA

  GPADataSheet.getRange("C2").setValue(roundedGPA); // Total GPA
  GPADataSheet.getRange("C3").setValue(schoolOneRoundedGPA); // School 1 GPA
  GPADataSheet.getRange("C4").setValue(schoolTwoRoundedGPA); // School 2 GPA

  GPADataSheet.getRange("D2").setValue(totalCreditHours); // Total Credit Hours
  GPADataSheet.getRange("D3").setValue(schoolOneCreditHours); // School One Credit Hours
  GPADataSheet.getRange("D4").setValue(schoolTwoCreditHours); // School Two Credit Hours

  GPADataSheet.getRange("G2").setValue(totalWeightedPoints);
  GPADataSheet.getRange("G3").setValue(schoolOneWeightedPoints);
  GPADataSheet.getRange("G4").setValue(schoolTwoWeightedPoints);

  GPADataSheet.getRange("F2").setValue(totalCreditsInProgress);
  GPADataSheet.getRange("F3").setValue(schoolOneCreditsInProgress);
  GPADataSheet.getRange("F4").setValue(schoolTwoCreditsInProgress);

  GPADataSheet.getRange("E2").setValue(totalEarnedCredits);
  GPADataSheet.getRange("E3").setValue(schoolOneEarnedCredits);
  GPADataSheet.getRange("E4").setValue(schoolTwoEarnedCredits);


  var numRows = sheet.getLastRow();
  var simplifiedLastRow = simplifiedSheet.getLastRow();
  if (simplifiedLastRow < numRows) {
    simplifiedSheet.insertRowsAfter(simplifiedLastRow, numRows - simplifiedLastRow);
  } // End if


  // Copy data over from Class Records to Simplified Records
  for (var i = 1; i <= numRows; i++) {
    if (simplifiedSheet.getRange(i, 5).getValue() !== sheet.getRange(i, 9).getValue()) {
      //console.log("Empty");
      simplifiedSheet.getRange(i, 1).setValue(sheet.getRange(i, 1).getValue()); // Semester
      simplifiedSheet.getRange(i, 2).setValue(sheet.getRange(i, 2).getValue()); // School
      simplifiedSheet.getRange(i, 3).setValue(sheet.getRange(i, 3).getValue()); // Course Name
      simplifiedSheet.getRange(i, 4).setValue(sheet.getRange(i, 8).getValue()); // Credit Hours
      simplifiedSheet.getRange(i, 5).setValue(sheet.getRange(i, 9).getValue()); // Grade Recieved
      simplifiedSheet.getRange(i, 6).setValue(sheet.getRange(i, 10).getValue()); // Grade Points
      simplifiedSheet.getRange(i, 7).setValue(sheet.getRange(i, 11).getValue()); // Weighted Points
    } else {
      //console.log("Not Empty");
    } // End if else
  } // End for

  simplifiedSheet.getRange("I2").setValue(sheet.getRange("M2").getValue()); // Credit Hours
  simplifiedSheet.getRange("J2").setValue(sheet.getRange("N2").getValue()); // Weighted Points
  simplifiedSheet.getRange("I5").setValue(sheet.getRange("M5").getValue()); // Total GPA
  simplifiedSheet.getRange("J5").setValue(sheet.getRange("N5").getValue()); // Rounded GPA
  simplifiedSheet.getRange("I8").setValue(sheet.getRange("M8").getValue()); // Total Earned Credits
  simplifiedSheet.getRange("J8").setValue(sheet.getRange("N8").getValue()); // Credits In Progress
  simplifiedSheet.getRange("I11").setValue(sheet.getRange("M11").getValue()); // Credits and Courses in the future
  simplifiedSheet.getRange("J11").setValue(sheet.getRange("N11").getValue()); // Credits and Courses in the planned future


} // End function calculateGPA

/**
 * Looks up grade points based on grade received using GradeVariables table.
 */
function lookupGradePoints(grade, gradeData) {
  for (var i = 1; i < gradeData.length; i++) {
    if (gradeData[i][0] == grade) {
      return parseFloat(gradeData[i][1]); // Return corresponding Grade Points
    } // End if
  } // End for
  return 0; // Default to 0 if grade not found
} // End function lookupGradePoints
