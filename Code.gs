/**
 * Copyright (c) 2021 Justin Yao Du. Licensed under the MIT License.
 * For more information, please visit:
 * https://github.com/justinyaodu/gsheet-planner
 */

/**
 * @OnlyCurrentDoc
 */

const START_TIME = Date.now();

const STATUS_NAMED_RANGE = "PlannerStatus";
const CONFIG_NAMED_RANGE = "PlannerConfig";

let spreadsheet;

/**
 * Event handler for edits.
 * @param {Event} e The edit event.
 */
function onEdit(e) {
  log("Entering onEdit");

  spreadsheet = e.source;

  Config.loadBase();
  log("Base config loaded");

  Status.clear();
  log("Status cleared");


  try {
    Config.loadUser();
    log("User config loaded");

    onEditExpandDateMacros(e);
    log("Date macros expanded");

    onEditSortPlanner(e);
    log("Planner sorted");
  } catch (err) {
    Status.error(err.toString());
  }

  log("Exiting onEdit");
}

function log(message) {
  console.log(`${Math.round(Date.now() - START_TIME)} ms: ${message}`);
}

/**
 * Event hook which expands date macros.
 * @param {Event} e The edit event.
 */
function onEditExpandDateMacros(e) {
  const modifiedDateRange = RangeUtils.intersection(e.range, Config.dateRange);
  if (modifiedDateRange === null) {
    return;
  }

  for (let row = 1; row <= modifiedDateRange.getHeight(); row++) {
    for (let column = 1; column <= modifiedDateRange.getWidth(); column++) {
      const cell = modifiedDateRange.getCell(row, column);
      const value = cell.getValue();

      let newValue;
      try {
        newValue = DateMacro.evaluate(value);
      } catch (err) {
        newValue = err.message;
      }

      if (newValue !== null) {
        cell.setValue(newValue);
      }
    }
  }
}

/**
 * Event hook which sorts the planner.
 * @param {Event} e The edit event.
 */
function onEditSortPlanner(e) {
  if (!Config.autoSort) {
    return;
  }

  // Resort if the planner or config changed. (The latter enables resorting whenever the autoSort
  // option is turned on, without having to make an extra edit in the planner.)
  if (RangeUtils.intersection(e.range, Config.plannerRange) !== null
      || RangeUtils.intersection(e.range, Config.configRange) !== null) {
    Config.plannerRange.sort(Config.plannerRange.getLastColumn());
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

/**
 * Manage the status cell.
 */
class Status {
  /**
   * Display an error message.
   * @param {string} message The error message.
   */
  static error(message) {
    const text = `${message} (${new Date()})`;
    Config.statusRange.setValue(text);
  }

  /**
   * Clear the status.
   */
  static clear() {
    RangeUtils.fastClearContent(Config.statusRange);
  }
}

/**
 * Load configuration attributes onto this object.
 */
class Config {
  /**
   * Load the non-configurable attributes.
   * @throws {Error} If the hardcoded named ranges do not exist.
   */
  static loadBase() {
    Config.configRange = RangeUtils.getNamedRange(CONFIG_NAMED_RANGE);
    Config.statusRange = RangeUtils.getNamedRange(STATUS_NAMED_RANGE);
  }

  /**
   * Load the user's config from the named range given by CONFIG_NAMED_RANGE.
   * @throws {Error} If the user's configuration has errors.
   */
  static loadUser() {
    const range = Config.configRange;
    if (range.getWidth() !== 3) {
      throw new Error(`Config range has ${range.getWidth()} columns, expected 3`);
    }

    // Clear the status column.
    RangeUtils.fastClearContent(range.offset(0, range.getWidth() - 1, range.getHeight(), 1));

    for (let row = 1; row <= range.getHeight(); row++) {
      Config.processConfigRow(row);
    }

    Config.assertAllDefined();
  }

  /**
   * Update the config by processing a row as a key value pair.
   * @param {number} row The row number.
   * @throws {Error} If the row failed to process.
   */
  static processConfigRow(row) {
    const range = Config.configRange;
    const key = range.getCell(row, 1).getValue();
    const value = range.getCell(row, 2).getValue();
    const statusCell = range.getCell(row, 3);

    if (key === "") {
      return;
    }

    try {
      const transformer = ConfigTransformers[key];
      if (transformer === undefined) {
        throw new Error(`Unrecognized key '${key}'`);
      }

      Config[key] = ConfigTransformers[key](value);
    } catch (err) {
      statusCell.setValue(err.toString());
      throw new Error("Config error");
    }
  }

  /**
   * Assert that every config attribute is defined.
   * @throws {Error} If any config attributes are missing.
   */
  static assertAllDefined() {
    for (const key of Object.getOwnPropertyNames(ConfigTransformers)) {
      if (Config[key] === undefined) {
        throw new Error(`Config attribute '${key}' not defined`);
      }
    }
  }
}

/**
 * Validate and transform config attributes. Each function in this class is named after the
 * config attribute it manages. The functions map the value in the config row to the value of the
 * corresponding attribute in the config object. They throw exceptions if the value is invalid.
 */
class ConfigTransformers {
  static autoSort(value) {
    assertType(value, "boolean");
    return value;
  }

  static plannerRange(value) {
    assertType(value, "string");
    return spreadsheet.getRange(value);
  }

  static dateRange(value) {
    assertType(value, "string");
    return spreadsheet.getRange(value);
  }
}

/**
 * Assert that a value has the given type.
 * @param {*} value The value to check.
 * @param {string} type The type to check for, as returned by typeof.
 * @throws {TypeError} If the types do not match.
 */
function assertType(value, type) {
  if (typeof value !== type) {
    throw TypeError(`Expected ${type}, got ${typeof value}`);
  }
}

/**
 * Range manipulation functions.
 */
class RangeUtils {
  /**
   * Return the intersection of two Ranges.
   * @param {Range} a The first Range.
   * @param {Range} b The second Range.
   * @returns {Range} The Ranges' intersection, or null if the Ranges do not intersect.
   */
  static intersection(a, b) {
    if (a.getSheet().getSheetId() !== b.getSheet().getSheetId()) {
      return null;
    }

    const rowInterval = RangeUtils.intervalIntersection(
      a.getRow(), a.getLastRow(), b.getRow(), b.getLastRow());

    const columnInterval = RangeUtils.intervalIntersection(
      a.getColumn(), a.getLastColumn(), b.getColumn(), b.getLastColumn());

    if (rowInterval === null || columnInterval === null) {
      return null;
    } else {
      const row = rowInterval[0];
      const numRows = rowInterval[1] - rowInterval[0] + 1;

      const column = columnInterval[0];
      const numColumns = columnInterval[1] - columnInterval[0] + 1;

      return a.getSheet().getRange(row, column, numRows, numColumns);
    }
  }

  /**
   * Return the intersection of two inclusive 1D intervals.
   * @param {number} a1 The left endpoint of the first interval.
   * @param {number} a2 The right endpoint of the first interval.
   * @param {number} b1 The left endpoint of the second interval.
   * @param {number} b2 The right endpoind of the second interval.
   * @returns {number[]} The left and right endpoints of the intersection interval, or null if the
   *     intervals do not intersect.
   */
  static intervalIntersection(a1, a2, b1, b2) {
    if (a1 > b1) {
      return RangeUtils.intervalIntersection(b1, b2, a1, a2);
    }

    if (a2 < b1) {
      return null;
    } else {
      return [b1, Math.min(a2, b2)];
    }
  }

  /**
   * Look up a named range by its name.
   * @param {string} name The name of the named range.
   * @returns {Range} The range of the named range.
   * @throws {Error} If there is no named range with the specified name.
   */
  static getNamedRange(name) {
    const matchingRanges = spreadsheet.getNamedRanges()
      .filter((namedRange) => namedRange.getName() === name);

    if (matchingRanges.length === 0) {
      throw new Error(`Named range '${name}' does not exist`);
    }

    return matchingRanges[0].getRange();
  }

  /**
   * Clear the content of a range, unless it is already blank.
   * @param {range} The range to clear.
   */
  static fastClearContent(range) {
    if (!range.isBlank()) {
      range.clearContent();
    }
  }
}
