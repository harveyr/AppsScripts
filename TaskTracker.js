/**
 * Add some functionality to help me use a Google Sheet as my task manager.
 *
 * Highly experimental, but looks promising. This was pretty easy and fun to
 * throw together!
 *
 * This is a "bound script" — it's used just within the one sheet.
 * Docs: https://developers.google.com/apps-script/guides/bound
 *
 * Source: https://github.com/harveyr/AppsScripts/blob/main/TaskTracker.js
 */

const CHECK_COMPLETE = "✔️"

const TASK_NOTES_FOLDER_ID = "11761DZMMmHJq4vYh-fB5JYGYNI8WPaAz"

/**
 * Set up the custom menu on open.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('Harvey')
    .addItem('🧹 Tidy Completed Tasks', 'tidyCompleted')
    .addItem('✍️ Create Task Doc', 'createTaskNote')
    .addToUi()
}

/**
 * On edit, check if we marked the task complete. If so, annotate with a
 * completed-on date.
 */
function onEdit(e) {
  // TBD: This one didn't behave the same way:
  // const sheet = e.source

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  // Get the selected cell.
  const range = e.range
  const row = range.getRow()
  const col = range.getColumn()

  // Find the column for checkmarks. If that's not the column we edited, bail.
  const checkColNum = findHeaderColumnNumber(sheet, CHECK_COMPLETE)
  if (col !== checkColNum) {
    return
  }

  // Find the column where we annotate with the completion date.
  const completedColNum = findHeaderColumnNumber(sheet, 'Completed')

  // Determine what edit was made.
  const newVal = range.getValue() || ''
  const oldVal = e.oldValue || ''
  const wasCheckedComplete = oldVal === '' && newVal === CHECK_COMPLETE;

  // If the task checked complete, annotate with today's date. Otherwise, clear
  // the annotation.
  let completedVal = ''
  if (wasCheckedComplete) {
    completedVal = new Date().toDateString()
  }

  // Set the annotation.
  sheet.getRange(row, completedColNum).setValue(completedVal)
}

/**
 * Utility function to find a column by its header name.
 *
 * @param {Sheet} sheet The current sheet. See
 * https://developers.google.com/apps-script/reference/spreadsheet/sheet.
 * @param {string} name The header name.
 * @returns number
 */
function findHeaderColumnNumber(sheet, name) {
  const headerNames = sheet.getRange(`A1:Z1`).getValues()[0]
  const linkColIdx = headerNames.findIndex(obj => {
    return obj === name
  })

  if (linkColIdx === -1) {
    throw new Error(`Couldn't find header with name: ${name}`)
  }

  return linkColIdx + 1;
}

/**
 * Menu handler to create a Google Doc for the current task. These docs are
 * where I take verbose notes on the task, write drafts, etc.
 */
function createTaskNote() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  const curCell = sheet.getActiveCell()
  const curRow = curCell.getRow()

  // Find the task name for the active row.
  const taskName = sheet.getRange(`A${curRow}`).getValue()
  if (!taskName) {
    SpreadsheetApp.getUi().alert(`No task name found for row ${curRow}`)
    return
  }

  // Find the "Link" column. This is where I keep the link to the doc.
  // TODO: use my new utility function instead.
  const headerNames = sheet.getRange(`A1:Z1`).getValues()[0]
  const linkColNum = headerNames.findIndex(obj => {
    return obj === "Link"
  }) + 1

  // Check if there's anything already in the link cell. If so, bail for now.
  const targetRange = sheet.getRange(curRow, linkColNum)
  if (targetRange.getValue().toString().length) {
    SpreadsheetApp.getUi().alert(`There's already data in ${targetRange.getA1Notation()}`)
    return
  }

  // Create the doc.
  const doc = DocumentApp.create(`[task note] ${taskName}`)
  const docUrl = doc.getUrl()

  // Move the doc to my folder.
  // See https://stackoverflow.com/a/31754059
  const docFile = DriveApp.getFileById(doc.getId())
  DriveApp.getFolderById(TASK_NOTES_FOLDER_ID).addFile( docFile )
  DriveApp.getRootFolder().removeFile(docFile)

  // Add the link to the sheet.
  targetRange.setValue(docUrl)
}

/**
 * Menu item handler to tidy up the list by moving completed items down under
 * the "Done" section.
 */
function tidyCompleted() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  // Find row starting the "Done" section. This is down below my incomplete
  // tasks.
  const nameDat = sheet.getRange("A1:A100").getValues()
  const doneRowIndex = nameDat.findIndex(obj => {
    return obj[0] === "Done"
  })
  const doneRowNum = doneRowIndex + 1

  // Look throw all rows up until the "Done" section to find
  // completed-but-untidied tasks.
  const taskRange = `A1:B${doneRowNum}`
  const taskDat = sheet.getRange(taskRange).getValues()
  taskDat.forEach((obj, index) => {
    // Skip header row. I start at the header row, even though I know I don't
    // care about it, to keep the `index` -> `row number` math simpler.
    if (!index) {
      return
    }

    const rowNum = index + 1
    const name = obj[0]
    const isChecked = obj[1] === CHECK_COMPLETE

    if (!isChecked) {
      return
    }

    // Create a new row just after the "Done" header row to copy the values
    // into. Note that I use `insertRowBefore`, instead of `insertRowAfter`, so
    // that the new row takes on the styling of a task row instead of the "Done"
    // header row.
    const targetRowNum = doneRowNum + 1
    sheet.insertRowBefore(targetRowNum)

    // Copy the current row into the new destination row.
    const range = `A${rowNum}:Z${rowNum}`
    const row = sheet.getRange(range)
    row.copyTo(sheet.getRange(`A${targetRowNum}:Z${targetRowNum}`))

    // Delete the current row.
    sheet.deleteRow(rowNum)
  })
}
