
function updateDashboard()
{
  var schools = DIRECTORY_TABLE.getRows()
  var firstSchool = schools[0]
  if (!(firstSchool && firstSchool.schoolName && firstSchool.schoolName))
  {
    throw new Error("There are no schools listed in the School Directory tab");
  }
  showSpinnerModal(
    "Please wait while we compile STARR data for each school.",
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

/**
 * This is called by the spinner modal to initiate a batch of updateVendorShipping
 */
function importSchool(continuationArgs)
{
  console.log("Importing school with continuationArgs: %s", JSON.stringify(continuationArgs))
  // Get two schools, so that we know the name of the next school to pass back (and also whether we are at the end)
  var sheetRow = DIRECTORY_TABLE.defaultOptions.headersRowIndex + continuationArgs.continuationToken + 1
  var schools = DIRECTORY_TABLE.getRows(null,
    sheetRow + ":" + (sheetRow + 1)
  )
  var school = schools[0]
  console.log("Importing school %s", JSON.stringify(school))

  var lastError = null;

  try
  {
    processSchool(school)
  } catch (err)
  {
    notifyError(err, false, "Error updating dashboard for school " + JSON.stringify(school))
    lastError = "⚠️ There was an error importing data for " + school.schoolName
  }

  // If there is a next school, return its name, otherwise we are done
  if (schools[1])
  {
    return {
      continuationToken: continuationArgs.continuationToken + 1,
      currentSchool: schools[1].schoolName,
      lastError: lastError,
      schoolCount: continuationArgs.schoolCount
    }
  }
  else
  {
    return null
  }
};

function processSchool(school)
{
  var schoolTable = new SheetsTable({
    name: school.schoolName,
    spreadsheetUrl: school.starrSpreadsheetLink,
    sheetIndex: 0,
    primaryKey: null,
    defaultOptions: {
      headersRowIndex: 4,
      headersCase: 'camel',
    }
  });
  var headers = schoolTable.headers
  // Early participation squalls have a slightly different set of headers
  var participationType = headers.includes('lastName') ? 'standard' : 'early';
  // The header that identifies a row, allows us to process sub rows
  var idHeader = participationType === 'standard' ? 'lastName' : 'gradeBand';
  // Filter for rows with relevant data
  var schoolData = schoolTable.getRows()
    .filter(x =>
    {
      return x.actionSteps &&
        (participationType === 'standard' ?
          x.lastName !== '(Last Name)' :
          x.gradeBand !== 'Choose Grade Band')
    })
  console.log("Found %s data rows", schoolData.length)

  // Form a nested array of each teacher's sub rows
  var teachers = [];
  if (schoolData.length > 0)
  {
    var teacher;
    schoolData.forEach(row =>
    {
      if (row[idHeader])
      {
        if (teacher) teachers.push(new Object(teacher))
        teacher = [row]
      }
      else
      {
        teacher.push(row);
      }
    })
    teachers.push(teacher)
  }

  // Build the data object
  var schoolSummary = {
    'districtName': school.districtName,
    'schoolName': school.schoolName,
    'totalNumberOfTeachersleaders': teachers.length,
    // One goal per teacher
    'totalNumberOfGoals': teachers.length,
    // Focus standard is same for entire school
    'schoolsFocusStandard': teachers.length > 0 ? teachers[0][0].focusStandard : null,
    // This one is manually entered
    'schoolsPlStrategy': (schoolTable.sheet.getRange('D3').getValue() || '').toString().replace(/^PL Strategy:\s*/i, ''),
    'lastUpdated': new Date()
  }
  if (schoolData.length > 0)
  {
    addGoalDates(schoolSummary, schoolData)
    addGoalCounts(schoolSummary, schoolData)
    addProgressCounts(schoolSummary, schoolData)
    addGoalStatus(schoolSummary, teachers)
  }
  console.log("School summary is %s", JSON.stringify(schoolSummary))

  DASHBOARD_TABLE.updateRows([schoolSummary], null, null,
    {
      'upsert': true,
      'onlyPresentColumns': true
    }
  )
}

function addGoalDates(schoolSummary, rows)
{
  var startDates = rows.map(row => { return row.smartGoalStartDate }).filter(x => { return x instanceof Date });
  schoolSummary.goalStartDate = startDates.length === 0 ? 'N/A' : new Date(Math.max(...startDates))
  var endDates = rows.map(row => { return row.smartGoalCompletedDate }).filter(x => { return x instanceof Date });
  schoolSummary.goalEndDate = endDates.length === 0 ? 'N/A' : new Date(Math.max(...endDates))
}

function addGoalCounts(schoolSummary, rows)
{
  var counts = countBy(rows, 'statusOfActionSteps')
  for (var prop in counts)
  {
    schoolSummary[normalizeHeader('Action Steps ' + prop)] = counts[prop];
  }
  schoolSummary.totalNumberOfActionSteps = counts.total
  schoolSummary.percentOfActionStepsCompleted = counts.total === 0 ? 0 : (counts['Completed'] || 0) / counts.total

  schoolSummary.totalNumberOfNotes = rows.filter(r => r.notes).length
}


function addProgressCounts(schoolSummary, rows)
{
  var counts = countBy(rows, 'progressIndicator')
  for (var prop in counts)
  {
    schoolSummary[normalizeHeader('Progress Indicator: ' + prop)] = counts[prop];
    schoolSummary[normalizeHeader('Percent of Progress Indicator: ' + prop)] = counts.total === 0 ? 0 : (counts[prop] || 0) / counts.total;
  }
}

function addGoalStatus(schoolSummary, teachers)
{
  schoolSummary.goalsMet = teachers.filter(t => t[0].statusOfGoal === 'Met').length
  schoolSummary.goalsNotMet = teachers.filter(t => t[0].statusOfGoal === 'Not Met').length
  schoolSummary.percentOfGoalsMet = schoolSummary.totalNumberOfGoals == 0 ? 0 : schoolSummary.goalsMet / schoolSummary.totalNumberOfGoals;
  schoolSummary.percentOfGoalsNotMet = schoolSummary.totalNumberOfGoals == 0 ? 0 : schoolSummary.goalsNotMet / schoolSummary.totalNumberOfGoals;
}
