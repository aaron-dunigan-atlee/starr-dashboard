
function updateDashboardWithCoaches() {

  // Get a list of coaches SPREADSHEETS
  // These will link to spreadhseets containing coach data
  var coachesSpreadsheets = COACHES_DIRECTORY_TABLE.getRows();

  // Check of any coach spreadsheets were found
  const [firstCoachSpreadsheet] = coachesSpreadsheets;
  if (!firstCoachSpreadsheet) {
    throw new Error("There are no coach spreadsheets listed in the Coaches Directory.");
  }

  // If coach spreadsheets were found, start iterating through them.
  showSpinnerModal(
    "Please wait while we compile STEARR data for each coach.",
    'importCoaches',
    '✔️ Dashboard updated',
    '⚠️ Something went wrong. Data may not have been updated.',
    "Updating Dashboard",
    {
      continuationToken: 0,
      currentSpreadsheet: firstCoachSpreadsheet.description,
      spreadsheetCount: coachesSpreadsheets.length
    }
  )

}

/**
 * This is called by the spinner modal to import data from one school
 */
function importCoaches(continuationArgs) {
  console.log("Importing coaches with continuationArgs: %s", JSON.stringify(continuationArgs))

  // Get two Coach Directory rows, so that we know the name of the next school to pass back (and also whether we are at the end)
  var sheetRow = COACHES_DIRECTORY_TABLE.defaultOptions.headersRowIndex + continuationArgs.continuationToken + 1
  var coachSpreadsheetRows = COACHES_DIRECTORY_TABLE.getRows(null,
    sheetRow + ":" + (sheetRow + 1)
  )

  var coachSpreadsheetRow = coachSpreadsheetRows[0]
  console.log("Importing coach spreadsheet %s", JSON.stringify(coachSpreadsheetRow))

  var lastError = null;

  try {
    processCoachSpreadsheet(coachSpreadsheetRow)
  } catch (err) {
    notifyError(err, false, "Error updating dashboard for coaches " + JSON.stringify(coachSpreadsheetRow))
    lastError = "⚠️ There was an error importing data for " + coachSpreadsheetRow.description
  }

  // If there is a next school, return its name, otherwise we are done
  if (coachSpreadsheetRows[1]) {
    return {
      continuationToken: continuationArgs.continuationToken + 1,
      currentSpreadsheet: coachSpreadsheetRows[1].description,
      lastError: lastError,
      spreadsheetCount: continuationArgs.spreadsheetCount
    }
  }
  else {
    return null
  }
};

function processCoachSpreadsheet(coachSpreadsheet) {

  var coachesTable = new SheetsTable({
    name: coachSpreadsheet.description,
    // The program was re-named, so we handle both versions
    spreadsheetUrl: coachSpreadsheet.coachesStearrSpreadsheetLink,
    sheetIndex: 0,
    primaryKey: null,
    defaultOptions: {
      headersRowIndex: 4,
      headersCase: 'camel',
    }
  });

  var headers = coachesTable.headers

  // The header that identifies a row, allows us to process sub rows
  var idHeader = 'lastName';

  // Filter for rows with relevant data
  var rawRows = coachesTable.getRows();
  var coachesData = rawRows.filter(x => {
    return x.actionSteps && x.lastName !== '(Last Name)';
  })

  //console.log("Found DATA", coachesData);

  // Form a nested array of each participant's sub rows
  // Participants are "practitioner" (teacher) or "leader"
  var participants = [];
  if (coachesData.length > 0) {
    var participant;
    coachesData.forEach((row, rowIndex) => {
      if (row[idHeader]) {
        if (participant) participants.push(new Object(participant))
        participant = [row]
      }
      else {
        try {
          participant.push(row);
        }
        catch (e) {
          console.error("This is probably a poorly formatted spreadsheet. Here's what we know:\n%s",
            JSON.stringify({
              idHeader,
              row,
              participants,
              rowIndex
            }, null, 2))
          throw e;
        }
      }
    })
    participants.push(participant)
  }

  console.log("PARTICIPANTS", participants);

  // Build the data object
  var coachRows = participants.map((participant, index) => {

    // Note, the coach contains the first action step
    // actionSteps contains the rest of the actions steps for that
    // coach in the form of [ { actionSteps: xxx, statusOfActionSteps: xxx } ]
    const [coach] = participant;

    const coachRow = {

      'districtName': coach.districtName,
      'schoolName': coach.schoolName,
      'coachName': coach.firstName + ' ' + coach.lastName,
      'coachschool': "Coach",
      'lastUpdated': new Date(),
      'schoolsFocusStandard': participants.length > 0 ? participants[0][0].focusStandard : null,
      'schoolsPlStrategy': (coachesTable.sheet.getRange('F3').getValue() || '').toString().replace(/^PL Strategy:\s*/i, ''),
      'totalNumberOfNotes': countCoachNotes(index, coachesTable),
      'goalStartDate': coach.smartGoalStartDate,
      'goalEndDate': coach.smartGoalCompletedDate,
      'coachesTotalNumberOfActionSteps': participant.length,
      'progressIndicatorModel': coach.overallProgress === 'Model' ? '✔️' : '-',
      'progressIndicatorImpact': coach.overallProgress === 'Impact' ? '✔️' : '-',
      'progressIndicatorInProgress': coach.overallProgress === 'In Progress' ? '✔️' : '-',
      'progressIndicatorCoaching': coach.overallProgress === 'Coaching' ? '✔️' : '-',
      'coachesGoalsMet': coach.statusOfCoachesActionPlanGoal,
      'leadersGoalsMet': "N/A",
      'practitionersGoalsMet': "N/A",
      'leadersPercentOfGoalsMet': "N/A",
      'leadersGoalsNotMet': "N/A",
      'practitionersGoalsNotMet': "N/A",
      'practitionersPercentOfGoalsMet': "N/A",
      'leadersPercentOfGoalsNotMet': "N/A",
      'practitionersPercentOfGoalsNotMet': "N/A",
      'statusOfCoachesActionPlanGoal': coach.statusOfCoachesActionPlanGoal,
      'reasonForNotMetCoach': coach.reasonForNotMet,
      'didCoachesImplementPlStrategy': participants[0][0].didCoachesImplementPlStrategy,
      'didSomeCoachesMeetTheirGoals': participants[0][0].didSomeCoachesMeetTheirGoals,
      'leadersPercentOfActionStepsCompleted': "N/A",
      'practitionersPercentOfActionStepsCompleted': "N/A",
      'leadersPercentOfProgressIndicatorModel': "N/A",
      'practitionersPercentOfProgressIndicatorModel': "N/A",
      'leadersPercentOfProgressIndicatorImpact': "N/A",
      'practitionersPercentOfProgressIndicatorImpact': "N/A",
      'leadersPercentOfProgressIndicatorInProgress': "N/A",
      'practitionersPercentOfProgressIndicatorInProgress': "N/A",
      'leadersPercentOfProgressIndicatorCoaching': "N/A",
      'practitionersPercentOfProgressIndicatorCoaching': "N/A"

    };

    if (coachesData.length > 0) {
      addCoachActionSteps(coachRow, participant);
      addCoachActionStepsCompletedPercentage(coachRow, participant);
    }

    return coachRow;

  });

  console.log("Coach rows are %s", JSON.stringify(coachRows))

  if (coachRows.length > 0) {
    DASHBOARD_TABLE.updateRows(coachRows, null, null,
      {
        'upsert': true,
        'onlyPresentColumns': true
      }
    )
  }
}

function addCoachActionSteps(coachRow, participant) {

  const actionSteps = {};

  participant.forEach((row) => {
    if (!actionSteps[row.statusOfActionSteps]) actionSteps[row.statusOfActionSteps] = 1;
    else actionSteps[row.statusOfActionSteps] += 1;
  });

  coachRow['actionStepsCompleted'] = actionSteps['Completed'] || 0;
  coachRow['actionStepsInProgress'] = actionSteps['In Progress'] || 0;
  coachRow['actionStepsUpcoming'] = actionSteps['Upcoming'] || 0;

}

function addCoachActionStepsCompletedPercentage(coachRow, participant) {

  coachRow.coachesPercentOfActionStepsCompleted = participant.reduce((acc, row) => {
    if (row.statusOfActionSteps === 'Completed') return acc + 1;
    return acc;
  }, 0) / participant.length;

}

function countCoachNotes(coachIndex, coachesTable) {
  const sheet = SpreadsheetApp.openByUrl(coachesTable.spreadsheetUrl).getSheets()[coachesTable.sheetIndex]
  let rows = sheet.getDataRange().getValues()
    // Remove the first row and use the second row as headers
    .slice(1)
  const headers = rows.shift()
  // Remove two more header rows
  rows.shift()
  rows.shift()
  const notesColumn = headers.map((header, index) =>
    /Classroom or Conference Notes/i.test(header) ? index : null)
    .filter(x => x !== null)[0]
  // Today there are 11 columns for notes; tomorrow, who knows?
  const notesColumns = Array(11).fill().map((element, index) => index + notesColumn)
  console.log("Notes headers found on columns %s", JSON.stringify(notesColumns))
  return notesColumns.filter(x => rows[coachIndex * 5][x] && !(/insert note/i.test(rows[coachIndex * 5][x]))).length
}
