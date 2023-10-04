
function testSheet()
{
  console.log(COACHES_DIRECTORY_TABLE.sheet)
}

function testImportSchool()
{
  var school =
  {
    "schoolName": "Baldwin Online Academy",
    "districtName": "Baldwin ",
    "starrFolder": "https://drive.google.com/drive/folders/1HIHvO5a80VZvsxaKQt97bhp-Hc7DNiEv?usp=sharing",
    "starrSpreadsheetLink": "https://docs.google.com/spreadsheets/d/1DolWv4hs2iblbTSqvWuN_TjtTaVpb3pg-oeoRrtg_Jw/edit#gid=0"
  }

  processSchool(school)
}

function testDeMac()
{
  console.log(new Date(Math.max(new Date(), new Date())));
}

function testRichText()
{
  var richText = SS.getSheetByName('School Directory').getRange('A55').getRichTextValue();
  console.log(richText)
  console.log(richText.getText())
}
