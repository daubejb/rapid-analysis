//MIT License
//
//Copyright (c) 2018 daubedesign
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.


var ADD_ON_TITLE = 'Rapid Analysis'
var dd = DocumentApp.getActiveDocument()
var body = dd.getBody()
var userProperties = PropertiesService.getUserProperties()
var primaryColor = userProperties.getProperty('primaryColor')
var secondaryColor = userProperties.getProperty('secondaryColor')
var logo = userProperties.getProperty('logo')
var logoWidth = userProperties.getProperty('logoWidth')
var logoHeight = userProperties.getProperty('logoHeight')

var CENTERED = {}
CENTERED[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER

/*********************************SETUP***********************************/

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
  .addItem('Launch', 'showSidebar')
  .addSeparator()
  .addItem('Configuration', 'showConfigMenu')
  .addToUi()
}

function onInstall(e) {
  onOpen(e)
  userProperties.setProperty('primaryColor', '#388E3C')
  userProperties.setProperty('secondaryColor', '#818181')
  userProperties.setProperty('logo', 'https://storage.googleapis.com/daube-design-assets.appspot.com/daubedesign.png')
  userProperties.setProperty('logoWidth', 28)
  userProperties.setProperty('logoHeight', 28)
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
  .setTitle(ADD_ON_TITLE)
  DocumentApp.getUi().showSidebar(ui)
}

function showConfigMenu() {
  var menu = HtmlService.createTemplateFromFile('menu').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  DocumentApp.getUi().showModelessDialog(menu, 'Rapid Analyst - Configuration Options');
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent()
}

function getConfiguration() {
  var configs = {
    primaryColorCode: primaryColor,
    secondaryColorCode: secondaryColor,
    logoUrl: logo,
    logoWidth: logoWidth,
    logoHeight: logoHeight
  }
  return configs
}

function updatePrimaryColor(color) {
  userProperties.setProperty('primaryColor', color)
  primaryColor = userProperties.getProperty('primaryColor')
}

function updateSecondaryColor(color) {
  userProperties.setProperty('secondaryColor', color)
  secondaryColor = userProperties.getProperty('secondaryColor')
}

function updateLogoWidth(width) {
  userProperties.setProperty('logoWidth', width)
  logoWidth = userProperties.getProperty('logoWidth')
}

function updateLogoHeight(height) {
  userProperties.setProperty('logoHeight', height)
  logoHeight = userProperties.getProperty('logoHeight')
}

function updateLogoUrl(url) {
  userProperties.setProperty('logoUrl', url)
  logo = userProperties.getProperty('logo')
}
/*****************************MAIN FUNCTIONS*****************************/

function insertTemplate(template, id) {
  id = id || 'id'
  switch (template) {
    case 'Executive Summary': createExecutiveSummary(); break;
    case 'Raci Matrix': createRaciMatrix(); break;
    case 'Meeting Notes': createMeetingNotes(id); break;
    case 'Problem Statement': insertProblemStatement(); break;
    case 'Objectives': insertObjectivesAndSuccessCriteria(); break;
    case 'User Stories Table': insertUserStoriesTable(); break;
  }
}

function insertProblemStatement() {
  var cursor = obtainCursorLocation()
  var ps1 = 'The Problem of ' 
  var ps2 = '<insert statement of problems> '
  var ps3 = 'affects '
  var ps4 = '<name affected people, organizations, or customer groups>. '
  var ps5 = 'The impact is '
  var ps6 = '<name the impact (i.e., poor decisions, cost overruns, erroneous information or processes, slow response time to customers, etc.)>. '
  var ps7 = 'A successful solution would '
  var ps8 = '<describe the solution>.'
  var p = cursor.getElement()
  
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  p.appendText(ps1).setBold(true)
  p.appendText(ps2).setBold(false)
  p.appendText(ps3).setBold(true)
  p.appendText(ps4).setBold(false)
  p.appendText(ps5).setBold(true)
  p.appendText(ps6).setBold(false)
  p.appendText(ps7).setBold(true)
  p.appendText(ps8).setBold(false)
  var text = p.editAsText()
  text.setForegroundColor('#000000')
  text.setItalic(false) 
}

function obtainCursorLocation() {
  var cur = dd.getCursor() 
  if (cur) {
    return cur
  } else {
    DocumentApp.getUi().alert('Cannot insert text here.')
  }
}

/****************************TEMPLATE TEXT*******************************/

function createMeetingNotes(id) {
  var event = getSpecificEvent(id)
  var timeZone = getUserTimeZone()
  var startTime = Utilities.formatDate(event.getStartTime(), timeZone, "hh:mm aaa")
  var endTime = Utilities.formatDate(event.getEndTime(), timeZone, "hh:mm aaa")
  insertTitle(event.getTitle())
  insertDate(event.getStartTime(), timeZone)
  insertStartEndTime(startTime, endTime)
  insertLocation(event.getLocation())
  body.appendHorizontalRule()
  insertDescription(event.getDescription())
  body.appendHorizontalRule()
  insertMeetingNotes()
  body.appendHorizontalRule()
  insertGuestList(event.getGuestList()) 
}

function createRaciMatrix() {
  insertH2('RACI Matrix')
  insertNormal('This table describes the project\'s deliverables and the associated roles that are associated with the following levels of responsibility:\n')
  body.appendListItem('\(R\)esponsible:  The actual person performing the work.').editAsText().setBold(0, 14, true)
  body.appendListItem('\(A\)ccountable:  The one ultimately answerable for the correct completion of the deliverable or task and who delegates the work to the responsible party.').editAsText().setBold(0, 14, true)
  body.appendListItem('\(C\)onsulted:  Those whose opinions are sought, typically subject matter experts, and with whom you have two-way communication.').editAsText().setBold(0, 12, true)
  body.appendListItem('\(I\)nformed:  Those who are kept up-to-date on progress.').editAsText().setBold(0, 12, true)
  insertH3('Matrix')
  var raci = [
    ['Deliverable or action', 'IT Project Sponsor', 'Project Manager', 'Business Analyst', 'Developer', 'Business Owner'],
    ['Project charter', 'Informed', 'Accountable / Responsible', 'Consulted', 'Consulted', 'Consulted'],
    ['Requirements', 'Informed', 'Consulted', 'Responsible / Accountable', 'Consulted', 'Consulted']
  ]
  var raciTable = body.appendTable(raci)
  raciTable.setBorderColor('#818181')
  raciTable.editAsText().setForegroundColor('#000000').setFontSize(10)
}

function createExecutiveSummary() {
  if (dd.getFooter()) {
    var footer = dd.getFooter()
  } else {
    var footer = dd.addFooter()
  }
  var footerLogoImage = UrlFetchApp.fetch(logo)
  footer.appendImage(footerLogoImage.getBlob()).setWidth(logoWidth).setHeight(logoHeight).getParent().setAttributes(CENTERED)
  insertTitle('Executive Summary')
  insertH2('Introduction & background')
  insertGuidance('This section summarizes the rationale for the new service or enhancement to an existing service.  Provide a general description of the history or situation that lead to the recognition that the service should be built or enhanced.')
  insertH2('Opportunity')
  insertGuidance('Describe the problem that is being solved.  Describe the business environment for the service.  This may include a brief comparative evaluation of existing services and potential solutions, indicating why the proposed service is attractive.  Identify the problems that cannot be solved without the service, and how the service fits in with market trends or strategic directions.')
  insertProblemStatement()
  body.appendPageBreak()
  insertObjectivesAndSuccessCriteria()
  insertH2('Vision of the service')
  insertGuidance('Write a concise statement that summarizes the purpose and intent of the new service or for enhancing an existing service and describes what the business environment will be like when it includes the new / updated service.  The vision statement should reflect a balanced view that will satisfy the needs of diverse customers as well as those of IT Services.  It may be somewhat idealistic; however, it should be grounded in the realities of our company, enterprise architecture, organizational strategic directions, and cost and resource limitations.')
  body.appendPageBreak()
  insertH2('Analysis models')
  insertGuidance('Include any relevant analysis models including but not limited to: use case diagram, activity diagram, state diagram.')
  body.appendPageBreak()
  insertH2('Requirements')
  insertH3('Functional requirements')
  insertUserStoriesTable()
  body.appendPageBreak()
  insertH3('Nonfunctional requirements')
  insertH4('Performance')
  insertH4('Scalability')
  insertH4('Capacity')
  insertH4('Availability')
  insertH4('Reliability')
  insertH4('Recoverability')
  insertH4('Maintainability')
  insertH4('Serviceability')
  insertH4('Usability')
  insertH4('Security')
  insertH4('Regulatory and compliance')
  insertH4('Data integrity')
  insertH4('Interoperability')
  body.appendPageBreak()
  insertH2('Supplemental specification\(s\)')
  insertGuidance('Include any supplemental specifications not defined in or too lengthy to be described in the requirements above')
  insertH3('Supplemental specification #1')
  insertGuidance('This supplemental specification describes...')
  body.appendPageBreak()
  createRaciMatrix()
}

/****************************BUILDING BLOCKS*******************************/

function insertTitle(title) {
  title = title.toUpperCase()
  var t = body.appendParagraph(title)
  t.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  t.editAsText().setForegroundColor(primaryColor)
}

function insertH2(header) {
  var t = body.appendParagraph(header)
  t.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  t.editAsText().setForegroundColor(secondaryColor)
}

function insertH3(header) {
  var t = body.appendParagraph(header)
  t.editAsText().setForegroundColor('#000000')
  t.setHeading(DocumentApp.ParagraphHeading.HEADING3)
}

function insertH4(header) {
  var t = body.appendParagraph(header)
  t.setHeading(DocumentApp.ParagraphHeading.HEADING4)
}

function insertGuidance(text) {
  var p = body.appendParagraph(text)
  p.setItalic(true)
  p.editAsText().setForegroundColor('#818181') 
}

function insertNormal(text) {
  var p = body.appendParagraph(text)
  formatAsNormal(p)
}

function insertDate(date, timeZone) {
  var formattedDate = Utilities.formatDate(date, timeZone, "E, MMMM d,  y")
  var p = body.appendParagraph(formattedDate)
  formatAsNormal(p)
}

function insertStartEndTime(startTime, endTime) {
  var p = body.appendParagraph(startTime + ' - ' + endTime)
  p.editAsText().setForegroundColor('#818181')
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
}

function insertLocation(location) {
  location = 'Location:  ' + location
  var p = body.appendParagraph(location)
  p.editAsText().setForegroundColor('#818181')
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  p.editAsText().setBold(0, 9, true)
}

function insertDescription(description) {
  description = '\nDescription:\n\n' + description + '\n'
  var p = body.appendParagraph(description)
  formatAsNormal(p)
  p.editAsText().setBold(0, 12, true)
}

function insertMeetingNotes() {
  var p = body.appendParagraph('\nMeeting notes:\n\nTake notes here...\n')
  formatAsNormal(p)
  p.editAsText().setBold(0, 14, true)
}

function insertGuestList(guests) {
  var p = body.appendParagraph('\nMeeting invitees:\n')
  formatAsNormal(p)
  p.editAsText().setBold(0, 17, true)
  for (var i = 0, len = guests.length; i < len; i++) {
    body.appendParagraph(guests[i].getName() + '\(' + guests[i].getEmail() + '\): ' + guests[i].getGuestStatus())
  }
}

function insertProblemStatement() {
  var ps1 = 'The Problem of ' 
  var ps2 = '<insert statement of problems> '
  var ps3 = 'affects '
  var ps4 = '<name affected people, organizations, or customer groups>. '
  var ps5 = 'The impact is '
  var ps6 = '<name the impact (i.e., poor decisions, cost overruns, erroneous information or processes, slow response time to customers, etc.)>. '
  var ps7 = 'A successful solution would '
  var ps8 = '<describe the solution>.'
  var p = body.appendParagraph('\n')
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  p.appendText(ps1).setBold(true)
  p.appendText(ps2).setBold(false)
  p.appendText(ps3).setBold(true)
  p.appendText(ps4).setBold(false)
  p.appendText(ps5).setBold(true)
  p.appendText(ps6).setBold(false)
  p.appendText(ps7).setBold(true)
  p.appendText(ps8).setBold(false)
}

function insertObjectivesAndSuccessCriteria() {
  insertH2('Objectives and success criteria')
  var p = body.appendParagraph('Objectives include:\n')
  p.editAsText().setForegroundColor('#000000')
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
  body.appendListItem('<Improve business performance>')
  body.appendListItem('<Increase revenue')
  body.appendListItem('<Increase customer experience>')
  body.appendListItem('<Reduce cost>')
  body.appendListItem('<Mitigate risk / ensure compliance>')
  body.appendParagraph('\nSuccess of the service will be defined by:\n')
  var success = [
    ['Success criteria', 'Value provided', 'Measurement'],
    ['<Metric #1>', '<Value>', '<Objective measure>'],
    ['<Metric #2>', '<Value>', '<Objective measure>'],
    ['<Metric #3>', '<Value>', '<Objective measure>']
  ]
  var successTable = body.appendTable(success)
  successTable.setBorderColor('#818181')                               
}

function insertUserStoriesTable() {
  var userStories = [
    ['Role', 'Ability to', 'So that I can', 'Priority'],
    ['Account holder', 'Withdraw money from the ATM', 'Have quick access to my money all over the country', 'Must'],
    ['Account holder', 'Check my balance from the ATM', 'Check my account balance for reference', 'Must'],
    ['Technician', 'Check the amount of currency inside the ATM', 'Ensure that the ATM has money and my users can withdraw it', 'Nice to have']
  ]
  insertH4('User stories table')
  var userStoriesTable = body.appendTable(userStories)
  userStoriesTable.setBorderColor('#818181') 
}


/*********************************Calendar Functions********************************/

function getEventsFromToday() {
  var now = new Date()
  var ca = CalendarApp.getDefaultCalendar()
  var events = ca.getEventsForDay(now)
  var eventDetails = []
  var timeZone = getUserTimeZone()
  for (var i = 0, len = events.length; i < len; i++) {
    var n = events[i]
    var startTime = Utilities.formatDate(n.getStartTime(), timeZone, "hh:mm aaa")
    var endTime = Utilities.formatDate(n.getEndTime(), timeZone, "hh:mm aaa")
    eventDetails.push(
      {
        title: n.getTitle(),
        description: n.getDescription(),
        startTime: startTime,
        endTime: endTime,
        location: n.getLocation(),
        id: n.getId()
      }
    )
  }
  return JSON.stringify(eventDetails)
}

function getSpecificEvent(id) {
  var ca = CalendarApp.getEventById(id)
  return ca
}

function getUserTimeZone() {
  var userTimeZone = CalendarApp.getDefaultCalendar().getTimeZone();
  return userTimeZone
}

/*********************************Document Format Functions********************************/

function formatAsNormal(p) {
  p.editAsText().setForegroundColor('#000000')
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
}
