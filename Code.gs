/**
 * Copyright (c) 2021 Justin Yao Du. Licensed under the MIT License.
 * For more information, please visit:
 * https://github.com/justinyaodu/gsheet-planner
 */

/**
 * @OnlyCurrentDoc
 */

const START_TIME = Date.now();

const PLANNER_SHEET_NAME = "Planner";
const PLANNER_RANGE = "C2:K";
const DATE_COLUMN = 3;

/**
 * Log a message and the current execution time.
 * @param {string} message The message to log.
 */
function log(message) {
  console.log(`${Math.round(Date.now() - START_TIME)} ms: ${message}`);
}

/**
 * Event handler for edits.
 * @param {Event} e The edit event.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  log(`Got sheet '${sheet.getName()}'`);

  let plannerRange;

  if (sheet.getName() === PLANNER_SHEET_NAME) {
    plannerRange = sheet.getRange(PLANNER_RANGE);

    onEditExpandDateMacros(e);
    log("Date macros expanded");
  } else {
    plannerRange = sheet.getParent().getRange(PLANNER_SHEET_NAME + "!" + PLANNER_RANGE);
  }

  plannerRange.sort(plannerRange.getLastColumn());
  log("Planner sorted");
}

/**
 * Event hook which expands date macros.
 * @param {Event} e The edit event.
 */
function onEditExpandDateMacros(e) {
  if (DATE_COLUMN < e.range.getColumn() || e.range.getLastColumn() < DATE_COLUMN) {
    return;
  }

  let edited = false;

  const column = DATE_COLUMN - e.range.getColumn() + 1;
  for (let row = 1; row <= e.range.getHeight(); row++) {
    const cell = e.range.getCell(row, column);
    const value = cell.getValue();

    let newValue;
    try {
      newValue = DateMacro.evaluate(value);
    } catch (err) {
      newValue = err.message;
    }

    if (newValue !== null) {
      cell.setValue(newValue);
      edited = true;
    }
  }

  // Ensure that the Sort Date arrayformula is updated before sorting.
  if (edited) {
    SpreadsheetApp.flush();
  }
}

/**
 * Evaluate date macros into date strings. Each function with a single-letter name evaluates the
 * macros that start with that character.
 */
class DateMacro {
  /**
   * Evaluate a date macro.
   * @param {*} The value to evaluate.
   * @returns {string} The evaluated date, or null if the value is not a date macro.
   * @throws {Error} If the date macro could not be parsed.
   */
  static evaluate(value) {
    if (typeof value !== "string") {
      return null;
    }

    // Currently, all date macros start with a lowercase letter and an integer.
    const re = /^([a-z])(-?\d.*)$/;
    const result = re.exec(value);

    if (result === null) {
      return null;
    }

    const macroType = result[1];
    const macroText = result[2];

    const macroFunc = DateMacro[macroType];
    if (macroFunc === undefined) {
      throw new Error(`Date macro '${macroType}' not defined`);
    }

    return DateMacro.formatDate(macroFunc(macroText));
  }

  /**
   * Format a date.
   * @param {Date} date The date to format.
   * @returns {string} The date, formatted as YYYY-MM-DD.
   */
  static formatDate(date) {
    // Change the timezone to UTC while keeping the same date and time.
    const timezoneOffsetMs = date.getTimezoneOffset() * 60 * 1000;
    const utcDate = new Date(date - timezoneOffsetMs);

    return utcDate.toISOString().slice(0, 10);
  }

  /**
   * Parse a date from numbers separated by non-numeric characters. The numbers are interpreted as
   * date, month-date, or year-month-date, depending on how many numbers there are. If the month or
   * year are unspecified, the current month or year will be used.
   * @param {string} text 1-3 numbers, separated by non-numeric characters.
   * @returns {Date} The parsed date.
   * @throws {Error} If the date can't be parsed.
   */
  static d(text) {
    const re = /^(\d+)(?:\D+(\d+))?(?:\D+(\d+))?$/;
    const result = re.exec(text);

    if (result === null) {
      throw new Error(`Cannot parse '${text}' as date`);
    }

    // Get the matched regex groups in reverse order (day, month, year).
    const nums = result.slice(1).filter((num) => num !== undefined);
    nums.reverse();

    const date = new Date();

    const day = DateMacro.strictParseInt(nums[0]);

    const month = nums[1]
      ? DateMacro.strictParseInt(nums[1]) - 1
      : date.getMonth();

    const year = nums[2]
      ? DateMacro.parsePartialYear(nums[2])
      : date.getFullYear();

    date.setFullYear(year, month, day);
    return date;
  }

  /**
   * Parse 1-4 digits representing a year, inferring any missing digits. For example, if the current
   * year is 2021, "23" (or even "3") will become 2023.
   * @param {string} digits The digits to use as the suffix of the year number.
   * @returns {number} The parsed year number.
   * @throws {Error} If the suffix contains non-digit characters.
   */
  static parsePartialYear(digits) {
    const year = (new Date()).getFullYear().toString();
    const sliceEnd = year.length - digits.length;
    return DateMacro.strictParseInt(year.slice(0, sliceEnd) + digits);
  }

  /**
   * Return the current date plus the specified number of days.
   * @param {string} text The number of days to add. Negative values produce days in the past.
   * @returns {Date} The offset date.
   * @throws {Error} If the text can't be parsed as an integer.
   */
  static f(text) {
    const dateOffset = DateMacro.strictParseInt(text);

    const date = new Date();
    date.setDate(date.getDate() + dateOffset);
    return date;
  }

  /**
   * Return the next date with the specified weekday.
   * @param {string} text The weekday number (0 is Sunday, 1 is Monday, ..., 6 is Saturday).
   * @returns {Date} The next date with the specified weekday.
   * @throws {Error} If the text can't be parsed as an integer, or is out of range.
   */
  static w(text) {
    const weekday = DateMacro.strictParseInt(text);

    if (weekday < 0 || weekday >= 7) {
      throw new Error(`Weekday index '${weekday}' out of range`);
    }

    const date = new Date();

    do {
      date.setDate(date.getDate() + 1);
    } while (date.getDay() !== weekday);

    return date;
  }

  /**
   * Parse a string as an integer.
   * @param {string} str The string to parse.
   * @returns {number} The parsed integer.
   * @throws {Error} If the string cannot be parsed.
   */
  static strictParseInt(str) {
    const num = parseInt(str, 10);

    if (isNaN(num)) {
      throw new Error(`Cannot parse '${str}' as integer`);
    }

    return num;
  }
}
