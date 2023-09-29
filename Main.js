
function updateDashboard() {

  var schools = SCHOOLS_DIRECTORY_TABLE.getRows()
  var firstSchool = schools[0]
  if (!(firstSchool && firstSchool.schoolName && firstSchool.schoolName))
  {
    throw new Error("There are no schools listed in the School Directory tab");
  }
  showSpinnerModal(
    "Please wait while we compile STEARR data for each school.",
    'importSchool',
    '✔️ Dashboard updated',
    '⚠️ Something went wrong. Data may not have been updated.',
    "Updating Dashboard",
    {
      continuationToken: 0,
      currentSchool: firstSchool.schoolName,
      schoolCount: schools.length
    }
  )

}

function sortDashboard() {

  sortSheetByHeaders(
    DASHBOARD_TABLE.sheet,
    [
      { header: "Coach/School", ascending: false },
      { header: "School Name", ascending: true },
      { header: "District Name", ascending: true }
    ],
    { headersRowIndex: DASHBOARD_TABLE.defaultOptions.headersRowIndex }
  );

}

/**
 * This is called by the spinner modal to import data from one school
 */
function importSchool(continuationArgs) {
  console.log("Importing school with continuationArgs: %s", JSON.stringify(continuationArgs))
  // Get two schools, so that we know the name of the next school to pass back (and also whether we are at the end)
  var sheetRow = SCHOOLS_DIRECTORY_TABLE.defaultOptions.headersRowIndex + continuationArgs.continuationToken + 1
  var schools = SCHOOLS_DIRECTORY_TABLE.getRows(null,
    sheetRow + ":" + (sheetRow + 1)
  )
  var school = schools[0]
  console.log("Importing school %s", JSON.stringify(school))

  var lastError = null;

  try {
    processSchool(school)
  } catch (err) {
    notifyError(err, false, "Error updating dashboard for school " + JSON.stringify(school))
    lastError = "⚠️ There was an error importing data for " + school.schoolName
  }

  // If there is a next school, return its name, otherwise we are done
  if (schools[1]) {
    return {
      continuationToken: continuationArgs.continuationToken + 1,
      currentSchool: schools[1].schoolName,
      lastError: lastError,
      schoolCount: continuationArgs.schoolCount
    }
  }
  else {
    return null
  }
};

function processSchool(school) {
  var schoolTable = new SheetsTable({
    name: school.schoolName,
    // The program was re-named, so we handle both versions
    spreadsheetUrl: school.starrSpreadsheetLink || school.stearrSpreadsheetLink,
    sheetIndex: 0,
    primaryKey: null,
    defaultOptions: {
      headersRowIndex: 4,
      headersCase: 'camel',
    }
  });
  var headers = schoolTable.headers
  // Early participation schools have a slightly different set of headers
  var participationType = headers.includes('lastName') ? 'standard' : 'early';
  // The header that identifies a row, allows us to process sub rows
  var idHeader = participationType === 'standard' ? 'lastName' : 'gradeBand';
  // Filter for rows with relevant data
  var rawRows = schoolTable.getRows();
  var schoolData = rawRows.filter(x => {
    return x.actionSteps &&
      (participationType === 'standard' ?
        x.lastName !== '(Last Name)' :
        x.gradeBand !== 'Choose Grade Band')
  })
  console.log("Found %s data rows", schoolData.length)

  // Form a nested array of each participant's sub rows
  // Participants are "practitioner" (teacher) or "leader"
  var participants = [];
  if (schoolData.length > 0) {
    var participant;
    schoolData.forEach((row, rowIndex) => {
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

  // Build the data object
  let leaders = participants.filter(p => p[0].title === 'Leader')
  let practitioners = participants.filter(p => p[0].title === 'Practitioner')
  var schoolSummary = {
    'districtName': school.districtName,
    'schoolName': school.schoolName,
    'coachName': "N/A",
    'totalNumberOfLeaders': leaders.length,
    'totalNumberOfPractitioners': practitioners.length,
    // One goal per teacher
    'leadersTotalNumberOfGoals': leaders.length,
    'practitionersTotalNumberOfGoals': practitioners.length,
    // Focus standard is same for entire school
    'schoolsFocusStandard': participants.length > 0 ? participants[0][0].focusStandard : null,
    // This one is manually entered
    'schoolsPlStrategy': (schoolTable.sheet.getRange('D3').getValue() || '').toString().replace(/^PL Strategy:\s*/i, ''),
    'lastUpdated': new Date(),
    // 'totalNumberOfNotes': rawRows.filter(r => r.notes).length  // This was updated 8.23 to be multiple columns with month names
    'totalNumberOfNotes': countNotes(schoolTable),
    'coachschool': 'School'
  }
  if (schoolData.length > 0) {
    addGoalDates(schoolSummary, schoolData)
    addGoalCounts(schoolSummary, schoolData)
    addProgressCounts(schoolSummary, schoolData)
    addGoalStatus(schoolSummary, participants)
    Object.assign(schoolSummary, {
      statusOfSchoolActionPlanGoal: schoolData[0].statusOfSchoolActionPlanGoal,
      reasonForNotMet: schoolData[0].reasonForNotMet,
      didSchoolImplementPlStrategy: schoolData[0].didSchoolImplementPlStrategy,
      didSomeTeachersOrLeadersMeetTheirGoal: schoolData[0].didSomeTeachersOrLeadersMeetTheirGoal
    })
  }
  console.log("School summary is %s", JSON.stringify(schoolSummary))

  DASHBOARD_TABLE.updateRows([schoolSummary], null, null,
    {
      'upsert': true,
      'onlyPresentColumns': true
    }
  )
}


function countNotes(schoolTable) {
  const sheet = SpreadsheetApp.openByUrl(schoolTable.spreadsheetUrl).getSheets()[schoolTable.sheetIndex]
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
  return rows.reduce(function (acc, row) {
    let columns = notesColumns.filter(x => row[x] && !(/insert note/i.test(row[x])))
    console.log("Notes found on columns %s", JSON.stringify(columns));
    return acc
      + columns.length
  }, 0)
}

function addGoalDates(schoolSummary, rows) {
  var startDates = rows.map(row => { return row.smartGoalStartDate }).filter(x => { return x instanceof Date });
  schoolSummary.goalStartDate = startDates.length === 0 ? 'N/A' : new Date(Math.max(...startDates))
  var endDates = rows.map(row => { return row.smartGoalCompletedDate }).filter(x => { return x instanceof Date });
  schoolSummary.goalEndDate = endDates.length === 0 ? 'N/A' : new Date(Math.max(...endDates))
}

function addGoalCounts(schoolSummary, rows) {
  let leaders = rows.filter(r => r.title === 'Leader')
  let practitioners = rows.filter(r => r.title === 'Practitioner')
  var counts = countBy(rows, 'statusOfActionSteps')
  for (var prop in counts) {
    schoolSummary[normalizeHeader('Action Steps ' + prop)] = counts[prop];
  }
  schoolSummary.leadersTotalNumberOfActionSteps = leaders.length === 0 ? "N/A" : leaders.filter(x => x.statusOfActionSteps).length
  schoolSummary.practitionersTotalNumberOfActionSteps = practitioners.length === 0 ? "N/A" : practitioners.filter(x => x.statusOfActionSteps).length

  // Completion percent
  schoolSummary.leadersPercentOfActionStepsCompleted =
    schoolSummary.leadersTotalNumberOfActionSteps === 0 || schoolSummary.leadersTotalNumberOfActionSteps === "N/A" ? "N/A" :
      leaders.filter(x => x.statusOfActionSteps === 'Completed').length / schoolSummary.leadersTotalNumberOfActionSteps;
  schoolSummary.practitionersPercentOfActionStepsCompleted =
    schoolSummary.practitionersTotalNumberOfActionSteps === 0 || schoolSummary.practitionersTotalNumberOfActionSteps === "N/A" ? "N/A" :
      practitioners.filter(x => x.statusOfActionSteps === 'Completed').length / schoolSummary.practitionersTotalNumberOfActionSteps;
}


function addProgressCounts(schoolSummary, rows) {
  const progressIndicators = ['Model', 'Impact', 'In Progress', 'Coaching']
  var counts = countBy(rows, 'overallProgress') // 8.13.23 Header "progressIndicator" was changed to "overallProgress"
  progressIndicators.forEach(prop => {
    schoolSummary[normalizeHeader('Progress Indicator: ' + prop)] = counts[prop];
    // schoolSummary[normalizeHeader('Percent of Progress Indicator: ' + prop)] = counts.total === 0 ? 0 : (counts[prop] || 0) / counts.total;
  })
  let practitioners = rows.filter(x => x.title === 'Practitioner')
  let practitionerCounts = countBy(practitioners, 'overallProgress')
  // Use all progressIndicators, so we get 0's for empty keys
  progressIndicators.forEach(prop => {
    schoolSummary[normalizeHeader('Practitioners Percent of Progress Indicator: ' + prop)] =
      practitionerCounts.total === 0 ? "N/A" :
        (practitionerCounts[prop] || 0) / practitionerCounts.total;
  })
  let leaders = rows.filter(x => x.title === 'Leader')
  let leaderCounts = countBy(leaders, 'overallProgress')
  progressIndicators.forEach(prop => {
    schoolSummary[normalizeHeader('Leaders Percent of Progress Indicator: ' + prop)] =
      leaderCounts.total === 0 ? "N/A" :
        (leaderCounts[prop] || 0) / leaderCounts.total;
  })
}

function addGoalStatus(schoolSummary, participants) {
  // In the case where there are no goals to be met, print "N/A"
  schoolSummary.leadersGoalsMet = schoolSummary.leadersTotalNumberOfGoals == 0 ? "N/A" :
    participants.filter(t => t[0].statusOfGoal === 'Met' && t[0].title === 'Leader').length
  schoolSummary.practitionersGoalsMet = schoolSummary.practitionersTotalNumberOfGoals == 0 ? "N/A" :
    participants.filter(t => t[0].statusOfGoal === 'Met' && t[0].title === 'Practitioner').length
  schoolSummary.leadersGoalsNotMet = schoolSummary.leadersTotalNumberOfGoals == 0 ? "N/A" :
    participants.filter(t => t[0].statusOfGoal === 'Not Met' && t[0].title === 'Leader').length
  schoolSummary.practitionersGoalsNotMet = schoolSummary.practitionersTotalNumberOfGoals == 0 ? "N/A" :
    participants.filter(t => t[0].statusOfGoal === 'Not Met' && t[0].title === 'Practitioner').length

  schoolSummary.leadersPercentOfGoalsMet = schoolSummary.leadersTotalNumberOfGoals == 0 ? "N/A" :
    schoolSummary.leadersGoalsMet / schoolSummary.leadersTotalNumberOfGoals;
  schoolSummary.practitionersPercentOfGoalsMet = schoolSummary.practitionersTotalNumberOfGoals == 0 ? "N/A" :
    schoolSummary.practitionersGoalsMet / schoolSummary.practitionersTotalNumberOfGoals;
  schoolSummary.leadersPercentOfGoalsNotMet = schoolSummary.leadersTotalNumberOfGoals == 0 ? "N/A" :
    schoolSummary.leadersGoalsNotMet / schoolSummary.leadersTotalNumberOfGoals;
  schoolSummary.practitionersPercentOfGoalsNotMet = schoolSummary.practitionersTotalNumberOfGoals == 0 ? "N/A" :
    schoolSummary.practitionersGoalsNotMet / schoolSummary.practitionersTotalNumberOfGoals;
}
