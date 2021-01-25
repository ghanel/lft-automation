/**
 * Lateral flow testing AutoIT script generator
 * Generates AutoIT scripts for lateral flow tests.
 * Can be used for UK Government Website
 *
 * This script can be run on any computer however the resulting script will have to be run on a Windows computer with AutoIT installed. 
 * https://www.autoitscript.com/site/
 * 
 * The Windows computer should also have Google Chrome browser installed 
 * 
 * Items that you may wish to change are indicated by comments
 * 
 * 
 * 
 * 
 * 
 * ENSURE YOU CHANGE YOUR SITE ID!!!  Line 119
 * 
 * 
 * Current limitations (because we haven't found them necessary or government didn't ask us to collect the information)
 * Futher info on Ethnicity not input (only Asian/Black/White/Mixed/Other. No follow up for White of British/Irish etc). Left as "Prefer not to say"
 * "Prefer not to say" entered for travel to workplace question
 * Country is hard coded to England change 
 * Change line 270 to
 * Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}") for Scotland
 * Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{TAB}{ENTER}") for Northern Ireland
 * Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{DOWN}{TAB}{ENTER}") for Wales
 * 
 * Email address and mobile number are required
 * Landline number is not collected or input
 * NHS number is not collected or input
 * Vaccine question is hard coded to "No"
 * 
 * 
 * If you find this useful, please visit www.id3.org.uk and follow @id3company 
 * or follow me @gavinhanel
 *
 * @author Gavin Hanel https://www.id3.org.uk gavin@id3.org.uk  
 * @lastmodified 25 January 2021
 * @version 1 Gavin Hanel
 */

/* @license MIT License

Copyright (c) 2021 H&H Technology Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LFT Testing')
    .addItem('Generate files', 'run')
    .addToUi()
}

function run() {
  const ss = SpreadsheetApp.getActive()
  /**
   * 
   * Change the name of the Google Sheet here if necessary
   * 
   */
  const ws = ss.getSheetByName('Student Register')  //name of the Google Sheet
  let [headers, ...data] = ws.getDataRange().getValues()

  data.forEach((r, i) => {
  /**
   * 
   * Change the column heading from Barcode here if necessary (cell D1)
   * 
   */
    if (r[0] == true && r[headers.indexOf('Barcode')] != "") {
      generateScript(r, headers)
    }
  })

}


function generateFile(code, barcode) {
  /**
   * 
   * change this name to your liking.  It will be in your Google My Drive
   * 
   */
  let folderName = 'NHS LFT Registration'

  let folder = DriveApp.getFoldersByName(folderName).next() || DriveApp.createFolder(folderName)
  folder.createFile(barcode + ".au3", code, 'text/plain')
}

function generateScript(row, headers) {

  let output = ""

  /**
   * 
   * change your SITE ID
   * 
   */
  const siteID = 'CCCC'  //4 letter site code unique to your school

  
  
  /**
   * 
   * change any references you use in your spreadsheet here
   * e.g. for Male you may use an M
   * 
   */

  let sevenDayTestingTrue = "TRUE" 
  let sevenDayTestingFalse = "FALSE"
  let amIndicator = "AM"
  let pmIndicator = "PM"
  let maleIndicator = "Male"
  let femaleIndicator = "Female"
  let asian = "Asian or Asian British"
  let black = "Black, African, Black British or Caribbean"
  let mixed = "Mixed or multiple ethnic groups"
  let white = "White"
  let other = "Another ethnic group"
  let symptomsNo = "No"
  let symptomsYes = "Yes"

  /**
   * 
   * If you have changed any heading names in the spreadsheet change them here too
   * 
   */

  let barcode = row[headers.indexOf('Barcode')]
  let testDay = row[headers.indexOf('Test day')]
  let testMonth = row[headers.indexOf('Test month')]
  let testYear = row[headers.indexOf('Test year')]
  let firstname = row[headers.indexOf('Firstname')]
  let surname = row[headers.indexOf('Surname')]
  let dobDay = row[headers.indexOf('DOBDay')]
  let dobMonth = row[headers.indexOf('DOBMonth')]
  let dobYear = row[headers.indexOf('DOBYear')]
  let sevenDayTesting = row[headers.indexOf('Seven day testing')]
  let testHour = row[headers.indexOf('Test hour')]
  let testAMPM = row[headers.indexOf('Test AM/PM')]
  let emailAddress = row[headers.indexOf('Email address')]
  let gender = row[headers.indexOf('Gender')]
  let ethnicity = row[headers.indexOf('Ethnicity')]
  let postcode = row[headers.indexOf('Postcode')]
  let address = row[headers.indexOf('First line address')]
  let mobile = row[headers.indexOf('Mobile')]
  let symptoms = row[headers.indexOf('Symptoms')]
  let symptomsDay = row[headers.indexOf('Symptoms day')]
  let symptomsMonth = row[headers.indexOf('Symptoms month')]
  let symptomsYear = row[headers.indexOf('Symptoms year')]
  let emailArray = emailAddress.split('@')

  output += `Func _WinWaitActivate($title,$text,$timeout=0)
WinWait($title,$text,$timeout)
If Not WinActive($title,$text) Then WinActivate($title,$text)
WinWaitActive($title,$text,$timeout)
EndFunc

ShellExecute("chrome.exe", "/incognito /start-maximized https://test-for-coronavirus.service.gov.uk/register-kit/", "", "")
_WinWaitActivate("Register a test kit - GOV.UK - Google Chrome","Chrome Legacy Window")
Send("{TAB}{TAB}{TAB}{ENTER}{TAB}{TAB}{ENTER}{TAB}{TAB}{TAB}{ENTER}")
_WinWaitActivate("Who are you registering for? - GOV.UK - Google Chrome","Chrome Legacy Window")
Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}")
_WinWaitActivate("Scan the barcode - GOV.UK - Google Chrome","Chrome Legacy Window")
Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}${barcode}{TAB}${barcode}{TAB}{ENTER}")
 _WinWaitActivate("Where are you getting your test? - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}${siteID}{DOWN}{TAB}{ENTER}")
`

  if (sevenDayTesting == sevenDayTestingTrue) {
    output += `_WinWaitActivate("Are you doing 7-day repeat testing? - GOV.UK - Google Chrome","Chrome Legacy Window")
  Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}{ENTER}")
  `
  } else if (sevenDayTesting == sevenDayTestingFalse) {
    output += `_WinWaitActivate("Are you doing 7-day repeat testing? - GOV.UK - Google Chrome","Chrome Legacy Window")
  Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}")
  `
  } else {
    output += `_WinWaitActivate("Are you doing 7-day repeat testing? - GOV.UK - Google Chrome","Chrome Legacy Window")
  Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{TAB}{ENTER}")
  `
  }

  output += `_WinWaitActivate("Sample test date and time - GOV.UK - Google Chrome","Chrome Legacy Window")
  Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{TAB}${testDay}{TAB}${testMonth}{TAB}${testYear}{TAB}${testHour}{TAB}`

  if (testAMPM == amIndicator) {
    output += `{SPACE}{TAB}{ENTER}")
    `
  } else if (testAMPM == pmIndicator) {
    output += `{DOWN}{TAB}{ENTER}")
    `
  }
  output += `_WinWaitActivate("What’s your name? - GOV.UK - Google Chrome","Chrome Legacy Window")
  Send("{TAB}{TAB}{TAB}{TAB}${firstname}{TAB}${surname}{TAB}{ENTER}") 
 _WinWaitActivate("What's your date of birth? - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}${dobDay}{TAB}${dobMonth}{TAB}${dobYear}{TAB}{ENTER}")
  `

  if (gender == maleIndicator) {
    output += `_WinWaitActivate("What’s your gender? - GOV.UK - Google Chrome","Chrome Legacy Window")
   Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}{ENTER}")
   `
  } else if (gender == femaleIndicator) {
    output += `_WinWaitActivate("What’s your gender? - GOV.UK - Google Chrome","Chrome Legacy Window")
        Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}")
   `
  }

  output += `_WinWaitActivate("What’s your ethnicity? - GOV.UK - Google Chrome","Chrome Legacy Window")
  `

  if (ethnicity == asian) {
    output += `Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}{ENTER}{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
   `
  } else if (ethnicity == black) {
    output += `Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
   `
  } else if (ethnicity == mixed) {
    output += `Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{TAB}{ENTER}{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
   `
  } else if (ethnicity == white) {
    output += `Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{DOWN}{TAB}{ENTER}{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
   `
  } else if (ethnicity == other) {
    output += `Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}{DOWN}{DOWN}{TAB}{ENTER}{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
   `
  } else {
    output += `Send("{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
   `
  }
  output += `_WinWaitActivate("Do you travel to nursery, work or a place of education? - GOV.UK - Google Chrome","Chrome Legacy Window")
  Send("{TAB}{TAB}{TAB}{TAB}{UP}{TAB}{ENTER}")
  `

  if (symptoms == symptomsNo) {
    output += ` _WinWaitActivate("Symptomatic or asymptomatic - GOV.UK - Google Chrome","Chrome Legacy Window")
   Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}")
   `
  } else if (symptoms == symptomsYes) {
    output += ` _WinWaitActivate("Symptomatic or asymptomatic - GOV.UK - Google Chrome","Chrome Legacy Window")
   Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}{ENTER}")
   _WinWaitActivate("Date of onset - GOV.UK - Google Chrome","Chrome Legacy Window")
       Send("{TAB}{TAB}{TAB}{TAB}${symptomsDay}{TAB}${symptomsMonth}{TAB}${symptomsYear}{TAB}{ENTER}")
   `
  }

  output += `_WinWaitActivate("Country of residence - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}{ENTER}")
 _WinWaitActivate("What's your postcode? - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}${address}{TAB}${postcode}{TAB}{ENTER}")
 _WinWaitActivate("Email address - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}${emailArray[0]}{SHIFTDOWN}'{SHIFTUP}${emailArray[1]}{TAB}${emailArray[0]}{SHIFTDOWN}'{SHIFTUP}${emailArray[1]}{TAB}{ENTER}")
 _WinWaitActivate("What's your mobile number? - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}${mobile}{TAB}${mobile}{TAB}{ENTER}")
 _WinWaitActivate("Landline telephone number - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}")
 _WinWaitActivate("Do you know your NHS number? - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}{TAB}{DOWN}{TAB}{ENTER}")
 _WinWaitActivate("Have you had a coronavirus vaccine? - GOV.UK - Google Chrome","Chrome Legacy Window")
 Send("{TAB}{TAB}{TAB}{TAB}{SPACE}{TAB}{ENTER}")
`


  generateFile(output, barcode)
}
