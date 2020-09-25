/**
* @OnlyCurrentDoc
*/

const PLANNER_SHEET_NAME = "Planner";
const PLANNER_SORT_START_ROW = 2;
const PLANNER_SORT_START_COL = 3;

function onEdit(ev) {
  onEditExpandMacro(ev);
  onEditSortPlanner(ev);
}

/**
 * If a cell was edited and now contains a macro, expand it.
 */
function onEditExpandMacro(ev) {
  if (ev.value) {
    try {
      const expanded = expandMacro(ev.value);
      if (expanded !== null) {
        ev.range.setValue(expanded);
      }
    } catch (err) {
      ev.range.setValue(`Error expanding macro '${text}': ${err}`);
    }
  }
}

/**
 * If the planner sheet was edited, resort it.
 */
function onEditSortPlanner(ev) {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() === PLANNER_SHEET_NAME) {
    sortPlanner(sheet);
  }
}

/**
 * Given the sheet of the planner, sort its contents.
 */
function sortPlanner(sheet) {
  const numRows = sheet.getLastRow() - PLANNER_SORT_START_ROW + 1;
  const numCols = sheet.getLastColumn() - PLANNER_SORT_START_COL + 1;

  const range = sheet.getRange(
    PLANNER_SORT_START_ROW, PLANNER_SORT_START_COL, numRows, numCols);

  // TODO should this be range.getLastColumn()?
  range.sort(sheet.getLastColumn());
}

/**
 * Parse a string to an integer, or throw a string with an error message.
 */
function strictParseInt(str) {
  const num = parseInt(str, 10);

  if (isNaN(num)) {
    throw `Not a number: '${str}'`;
  }

  return num;
}

/**
 * Evaluate the given text as an exclamation-mark macro. Return the expanded
 * macro, or null if the text is not a macro. Throw an error message if an
 * error occurs.
 */
function expandMacro(text) {
  if (text.length < 1 || text.charAt(0) !== '!') {
    return null;
  }
  
  const command = text.charAt(1);
  const arg = text.substring(2);
  
  switch (command) {
    case 'w':
      return dateString(nextDayWithWeekday(strictParseInt(arg)));
    case 'm':
      return dateString(dayThisMonth(strictParseInt(arg)));
    case 'f':
      return dateString(daysInFuture(strictParseInt(arg)));
    default:
      throw `character '${command}' does not correspond to any macro`;
  }
}

/**
 * Format a date as YYYY-MM-DD.
 */
function dateString(date) {
  return date.toISOString().substring(0, 10);
}

/**
 * Return the next date with the specified weekday, or throw an error message if
 * the weekday is out of range. 0 is Sunday, 1 is Monday, ..., 6 is Saturday.
 */
function nextDayWithWeekday(weekday) {
  if (weekday < 0 || weekday >= 7) {
    throw `weekday index '${weekday}' out of range`;
  }
  
  const date = new Date();
  
  do {
    date.setDate(date.getDate() + 1);
  } while (date.getDay() != weekday);
  
  return date;
}

/**
 * Return the current date, with the day of the month set to the provided value.
 */
function dayThisMonth(day) {
  const date = new Date();
  date.setDate(day);
  return date;
}

/**
 * Return the current date plus the given number of days.
 */
function daysInFuture(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return date;
}
