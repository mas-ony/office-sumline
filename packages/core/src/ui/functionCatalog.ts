/**
 * Function signatures and short descriptions for formula autocomplete, argument hints and the
 * Insert Function dialog. Functions the engine implements but this table does not describe get a
 * generic signature derived from their arity.
 */
import { registry } from '../core/formula/functions';

export interface FunctionInfo {
  name: string;
  category: string;
  /** Parameters; optional ones are wrapped in brackets, a trailing '…' marks repetition. */
  args: string[];
  description: string;
}

const CATEGORIES: Record<string, string> = {
  M: 'Math & trigonometry',
  S: 'Statistical',
  T: 'Text',
  L: 'Logical',
  R: 'Lookup & reference',
  D: 'Date & time',
  F: 'Financial',
  E: 'Engineering',
  I: 'Information',
  B: 'Database',
};

// NAME|category|arguments separated by ';'|description
const TABLE = `
ABS|M|number|Absolute value of a number.
ACOS|M|number|Arccosine of a number, in radians.
ACOSH|M|number|Inverse hyperbolic cosine.
ADDRESS|R|row_num;column_num;[abs_num];[a1];[sheet_text]|Builds a cell reference as text.
AGGREGATE|M|function_num;options;ref1;[ref2];…|Aggregate (sum, average, …) that can skip hidden rows and errors.
AND|L|logical1;[logical2];…|TRUE when every argument is TRUE.
ARABIC|M|text|Converts a Roman numeral to a number.
ASIN|M|number|Arcsine of a number, in radians.
ATAN|M|number|Arctangent of a number, in radians.
ATAN2|M|x_num;y_num|Arctangent of the point (x, y).
AVEDEV|S|number1;[number2];…|Average absolute deviation from the mean.
AVERAGE|S|number1;[number2];…|Arithmetic mean of the numbers.
AVERAGEA|S|value1;[value2];…|Mean that counts text as 0 and TRUE as 1.
AVERAGEIF|S|range;criteria;[average_range]|Mean of the cells that meet a condition.
AVERAGEIFS|S|average_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Mean of the cells that meet several conditions.
BASE|M|number;radix;[min_length]|Converts a number to text in another base.
BIN2DEC|E|number|Binary to decimal.
BIN2HEX|E|number;[places]|Binary to hexadecimal.
CEILING|M|number;significance|Rounds up to a multiple of significance.
CEILING.MATH|M|number;[significance];[mode]|Rounds up to a multiple of significance.
CELL|I|info_type;[reference]|Information about a cell's formatting, location or contents.
CHAR|T|number|Character for a code number.
CHOOSE|R|index_num;value1;[value2];…|Picks a value from the list by position.
CHOOSECOLS|R|array;col_num1;[col_num2];…|Returns the given columns of an array.
CHOOSEROWS|R|array;row_num1;[row_num2];…|Returns the given rows of an array.
CLEAN|T|text|Removes non-printable characters.
CODE|T|text|Code number of the first character.
COLUMN|R|[reference]|Column number of a reference.
COLUMNS|R|array|Number of columns in a reference or array.
COMBIN|M|number;number_chosen|Number of combinations.
CONCAT|T|text1;[text2];…|Joins texts and ranges.
CONCATENATE|T|text1;[text2];…|Joins texts.
CONVERT|E|number;from_unit;to_unit|Converts between measurement units.
CORREL|S|array1;array2|Correlation coefficient of two data sets.
COS|M|number|Cosine of an angle in radians.
COSH|M|number|Hyperbolic cosine.
COUNT|S|value1;[value2];…|Counts the numbers.
COUNTA|S|value1;[value2];…|Counts the non-empty values.
COUNTBLANK|S|range|Counts empty cells.
COUNTIF|S|range;criteria|Counts the cells that meet a condition.
COUNTIFS|S|criteria_range1;criteria1;[criteria_range2;criteria2];…|Counts the cells that meet several conditions.
COVARIANCE.P|S|array1;array2|Population covariance.
COVARIANCE.S|S|array1;array2|Sample covariance.
CUMIPMT|F|rate;nper;pv;start_period;end_period;type|Cumulative interest paid between two periods.
CUMPRINC|F|rate;nper;pv;start_period;end_period;type|Cumulative principal paid between two periods.
DATE|D|year;month;day|Serial number of a date.
DATEDIF|D|start_date;end_date;unit|Difference between two dates in days, months or years.
DATEVALUE|D|date_text|Converts date text to a serial number.
DAVERAGE|B|database;field;criteria|Average of matching database records.
DAY|D|serial_number|Day of the month (1-31).
DAYS|D|end_date;start_date|Number of days between two dates.
DAYS360|D|start_date;end_date;[method]|Days between dates on a 360-day year.
DB|F|cost;salvage;life;period;[month]|Fixed-declining balance depreciation.
DCOUNT|B|database;field;criteria|Counts numeric cells in matching records.
DCOUNTA|B|database;field;criteria|Counts non-empty cells in matching records.
DDB|F|cost;salvage;life;period;[factor]|Double-declining balance depreciation.
DEC2BIN|E|number;[places]|Decimal to binary.
DEC2HEX|E|number;[places]|Decimal to hexadecimal.
DECIMAL|M|text;radix|Converts text in another base to a number.
DEGREES|M|angle|Radians to degrees.
DGET|B|database;field;criteria|Single value from the matching record.
DMAX|B|database;field;criteria|Largest value in matching records.
DMIN|B|database;field;criteria|Smallest value in matching records.
DOLLAR|T|number;[decimals]|Formats a number as currency text.
DROP|R|array;rows;[columns]|Removes rows or columns from the start or end of an array.
DSUM|B|database;field;criteria|Sum of matching records.
EDATE|D|start_date;months|Date a number of months before or after a date.
EFFECT|F|nominal_rate;npery|Effective annual interest rate.
EOMONTH|D|start_date;months|Last day of the month, months away.
EVEN|M|number|Rounds up to the nearest even integer.
EXACT|T|text1;text2|TRUE when two texts are identical (case-sensitive).
EXP|M|number|e raised to a power.
EXPAND|R|array;rows;[columns];[pad_with]|Expands an array to the given size.
FACT|M|number|Factorial.
FALSE|L||The logical value FALSE.
FILTER|R|array;include;[if_empty]|Rows of an array that meet a condition.
FIND|T|find_text;within_text;[start_num]|Position of text within text (case-sensitive).
FIXED|T|number;[decimals];[no_commas]|Formats a number as text with fixed decimals.
FLOOR|M|number;significance|Rounds down to a multiple of significance.
FLOOR.MATH|M|number;[significance];[mode]|Rounds down to a multiple of significance.
FORECAST|S|x;known_ys;known_xs|Linear-trend prediction.
FORECAST.LINEAR|S|x;known_ys;known_xs|Linear-trend prediction.
FORMULATEXT|R|reference|Formula of a cell as text.
FREQUENCY|S|data_array;bins_array|Frequency distribution as a vertical array.
FV|F|rate;nper;pmt;[pv];[type]|Future value of an investment.
GCD|M|number1;[number2];…|Greatest common divisor.
GEOMEAN|S|number1;[number2];…|Geometric mean.
GROWTH|S|known_ys;[known_xs];[new_xs];[const]|Exponential-trend prediction.
HARMEAN|S|number1;[number2];…|Harmonic mean.
HEX2DEC|E|number|Hexadecimal to decimal.
HLOOKUP|R|lookup_value;table_array;row_index_num;[range_lookup]|Looks along the top row and returns a value from another row.
HOUR|D|serial_number|Hour (0-23).
HSTACK|R|array1;[array2];…|Stacks arrays side by side.
HYPERLINK|I|link_location;[friendly_name]|Creates a link.
IF|L|logical_test;[value_if_true];[value_if_false]|Returns one value when a condition is TRUE and another when it is FALSE.
IFERROR|L|value;value_if_error|Value, or an alternative when it is an error.
IFNA|L|value;value_if_na|Value, or an alternative when it is #N/A.
IFS|L|logical_test1;value_if_true1;[logical_test2;value_if_true2];…|Value of the first condition that is TRUE.
INDEX|R|array;row_num;[column_num];[area_num]|Value at a row and column of a range.
INDIRECT|R|ref_text;[a1]|Reference given as text.
INT|M|number|Rounds down to an integer.
INTERCEPT|S|known_ys;known_xs|Y-intercept of the regression line.
IPMT|F|rate;per;nper;pv;[fv];[type]|Interest part of a payment.
IRR|F|values;[guess]|Internal rate of return.
ISBLANK|I|value|TRUE when the cell is empty.
ISERR|I|value|TRUE for any error except #N/A.
ISERROR|I|value|TRUE for any error.
ISEVEN|I|number|TRUE for even numbers.
ISFORMULA|I|reference|TRUE when the cell holds a formula.
ISLOGICAL|I|value|TRUE for logical values.
ISNA|I|value|TRUE for #N/A.
ISNONTEXT|I|value|TRUE for anything that is not text.
ISNUMBER|I|value|TRUE for numbers.
ISODD|I|number|TRUE for odd numbers.
ISOWEEKNUM|D|date|ISO week number.
ISREF|I|value|TRUE for references.
ISTEXT|I|value|TRUE for text.
KURT|S|number1;[number2];…|Kurtosis.
LAMBDA|L|[parameter1];…;calculation|Creates a reusable function.
LARGE|S|array;k|k-th largest value.
LCM|M|number1;[number2];…|Least common multiple.
LEFT|T|text;[num_chars]|First characters of a text.
LEN|T|text|Number of characters.
LET|L|name1;value1;[name2;value2];…;calculation|Names intermediate results.
LINEST|S|known_ys;[known_xs];[const];[stats]|Linear regression statistics.
LN|M|number|Natural logarithm.
LOG|M|number;[base]|Logarithm in a base (10 by default).
LOG10|M|number|Base-10 logarithm.
LOOKUP|R|lookup_value;lookup_vector;[result_vector]|Looks up a value in a row or column.
LOWER|T|text|Converts text to lowercase.
MATCH|R|lookup_value;lookup_array;[match_type]|Position of a value in a range.
MAX|S|number1;[number2];…|Largest value.
MAXA|S|value1;[value2];…|Largest value, counting text and logicals.
MAXIFS|S|max_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Largest value among cells that meet conditions.
MDETERM|M|array|Matrix determinant.
MEDIAN|S|number1;[number2];…|Median.
MID|T|text;start_num;num_chars|Characters from the middle of a text.
MIN|S|number1;[number2];…|Smallest value.
MINA|S|value1;[value2];…|Smallest value, counting text and logicals.
MINIFS|S|min_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Smallest value among cells that meet conditions.
MINUTE|D|serial_number|Minute (0-59).
MINVERSE|M|array|Inverse matrix.
MIRR|F|values;finance_rate;reinvest_rate|Modified internal rate of return.
MMULT|M|array1;array2|Matrix product.
MOD|M|number;divisor|Remainder of a division.
MODE|S|number1;[number2];…|Most frequent value.
MODE.SNGL|S|number1;[number2];…|Most frequent value.
MODE.MULT|S|number1;[number2];…|All most-frequent values.
MONTH|D|serial_number|Month (1-12).
MROUND|M|number;multiple|Rounds to the nearest multiple.
N|I|value|Value converted to a number.
NA|I||The #N/A error.
NETWORKDAYS|D|start_date;end_date;[holidays]|Working days between two dates.
NETWORKDAYS.INTL|D|start_date;end_date;[weekend];[holidays]|Working days with custom weekends.
NOMINAL|F|effect_rate;npery|Nominal annual interest rate.
NORM.DIST|S|x;mean;standard_dev;cumulative|Normal distribution.
NORM.INV|S|probability;mean;standard_dev|Inverse normal distribution.
NORM.S.DIST|S|z;cumulative|Standard normal distribution.
NORM.S.INV|S|probability|Inverse standard normal distribution.
NOT|L|logical|Reverses a logical value.
NOW|D||Current date and time.
NPER|F|rate;pmt;pv;[fv];[type]|Number of payment periods.
NPV|F|rate;value1;[value2];…|Net present value.
NUMBERVALUE|T|text;[decimal_separator];[group_separator]|Converts text to a number, locale-independently.
ODD|M|number|Rounds up to the nearest odd integer.
OFFSET|R|reference;rows;cols;[height];[width]|Reference offset from a starting cell.
OR|L|logical1;[logical2];…|TRUE when any argument is TRUE.
PERCENTILE|S|array;k|k-th percentile.
PERCENTILE.EXC|S|array;k|k-th percentile, exclusive.
PERCENTILE.INC|S|array;k|k-th percentile, inclusive.
PERCENTRANK|S|array;x;[significance]|Rank of a value as a percentage.
PERMUT|M|number;number_chosen|Number of permutations.
PI|M||The number π.
PMT|F|rate;nper;pv;[fv];[type]|Periodic loan payment.
POWER|M|number;power|Number raised to a power.
PPMT|F|rate;per;nper;pv;[fv];[type]|Principal part of a payment.
PRODUCT|M|number1;[number2];…|Multiplies the numbers.
PROPER|T|text|Capitalizes each word.
PV|F|rate;nper;pmt;[fv];[type]|Present value of an investment.
QUARTILE|S|array;quart|Quartile of a data set.
QUARTILE.INC|S|array;quart|Quartile of a data set, inclusive.
QUOTIENT|M|numerator;denominator|Integer part of a division.
RADIANS|M|angle|Degrees to radians.
RAND|M||Random number between 0 and 1.
RANDARRAY|M|[rows];[columns];[min];[max];[whole_number]|Array of random numbers.
RANDBETWEEN|M|bottom;top|Random integer between two numbers.
RANK|S|number;ref;[order]|Rank of a number in a list.
RANK.AVG|S|number;ref;[order]|Rank, averaging ties.
RANK.EQ|S|number;ref;[order]|Rank, ties share the top rank.
RATE|F|nper;pmt;pv;[fv];[type];[guess]|Interest rate per period.
REGEXEXTRACT|T|text;pattern;[return_mode];[case_sensitivity]|Text that matches a regular expression.
REGEXREPLACE|T|text;pattern;replacement;[occurrence];[case_sensitivity]|Replaces text that matches a regular expression.
REGEXTEST|T|text;pattern;[case_sensitivity]|TRUE when text matches a regular expression.
REPLACE|T|old_text;start_num;num_chars;new_text|Replaces characters at a position.
REPT|T|text;number_times|Repeats text.
RIGHT|T|text;[num_chars]|Last characters of a text.
ROMAN|M|number;[form]|Converts a number to a Roman numeral.
ROUND|M|number;num_digits|Rounds to a number of digits.
ROUNDDOWN|M|number;num_digits|Rounds toward zero.
ROUNDUP|M|number;num_digits|Rounds away from zero.
ROW|R|[reference]|Row number of a reference.
ROWS|R|array|Number of rows in a reference or array.
RSQ|S|known_ys;known_xs|Square of the correlation coefficient.
SEARCH|T|find_text;within_text;[start_num]|Position of text within text (not case-sensitive, wildcards).
SECOND|D|serial_number|Second (0-59).
SEQUENCE|M|rows;[columns];[start];[step]|Array of sequential numbers.
SHEET|I|[value]|Sheet number.
SHEETS|I|[reference]|Number of sheets.
SIGN|M|number|Sign of a number (-1, 0 or 1).
SIN|M|number|Sine of an angle in radians.
SINH|M|number|Hyperbolic sine.
SKEW|S|number1;[number2];…|Skewness.
SLN|F|cost;salvage;life|Straight-line depreciation.
SLOPE|S|known_ys;known_xs|Slope of the regression line.
SMALL|S|array;k|k-th smallest value.
SORT|R|array;[sort_index];[sort_order];[by_col]|Sorts an array.
SORTBY|R|array;by_array1;[sort_order1];…|Sorts an array by other arrays.
SQRT|M|number|Square root.
STANDARDIZE|S|x;mean;standard_dev|Normalized value.
STDEV|S|number1;[number2];…|Sample standard deviation.
STDEV.P|S|number1;[number2];…|Population standard deviation.
STDEV.S|S|number1;[number2];…|Sample standard deviation.
STDEVA|S|value1;[value2];…|Sample standard deviation, counting text and logicals.
STDEVP|S|number1;[number2];…|Population standard deviation.
SUBSTITUTE|T|text;old_text;new_text;[instance_num]|Replaces occurrences of text.
SUBTOTAL|M|function_num;ref1;[ref2];…|Subtotal that ignores other subtotals.
SUM|M|number1;[number2];…|Adds the numbers.
SUMIF|M|range;criteria;[sum_range]|Sum of the cells that meet a condition.
SUMIFS|M|sum_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Sum of the cells that meet several conditions.
SUMPRODUCT|M|array1;[array2];…|Sum of the products of corresponding items.
SUMSQ|M|number1;[number2];…|Sum of squares.
SWITCH|L|expression;value1;result1;[value2;result2];…;[default]|Matches an expression against values.
SYD|F|cost;salvage;life;per|Sum-of-years' digits depreciation.
T|T|value|Text, or empty text for other values.
TAKE|R|array;rows;[columns]|Rows or columns from the start or end of an array.
TAN|M|number|Tangent of an angle in radians.
TANH|M|number|Hyperbolic tangent.
TEXT|T|value;format_text|Formats a number as text.
TEXTAFTER|T|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|Text after a delimiter.
TEXTBEFORE|T|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|Text before a delimiter.
TEXTJOIN|T|delimiter;ignore_empty;text1;[text2];…|Joins texts with a delimiter.
TEXTSPLIT|T|text;col_delimiter;[row_delimiter];[ignore_empty];[match_mode];[pad_with]|Splits text into an array.
TIME|D|hour;minute;second|Serial number of a time.
TIMEVALUE|D|time_text|Converts time text to a serial number.
TOCOL|R|array;[ignore];[scan_by_column]|Array as a single column.
TODAY|D||Current date.
TOROW|R|array;[ignore];[scan_by_column]|Array as a single row.
TRANSPOSE|R|array|Swaps rows and columns.
TREND|S|known_ys;[known_xs];[new_xs];[const]|Linear-trend values.
TRIM|T|text|Removes extra spaces.
TRIMMEAN|S|array;percent|Mean excluding a percentage of outliers.
TRUE|L||The logical value TRUE.
TRUNC|M|number;[num_digits]|Truncates to an integer or digits.
TYPE|I|value|Type of a value as a number.
UNICHAR|T|number|Unicode character for a code point.
UNICODE|T|text|Code point of the first character.
UNIQUE|R|array;[by_col];[exactly_once]|Unique rows or columns of an array.
UPPER|T|text|Converts text to uppercase.
VALUE|T|text|Converts text to a number.
VALUETOTEXT|T|value;[format]|Value as text.
VAR|S|number1;[number2];…|Sample variance.
VAR.P|S|number1;[number2];…|Population variance.
VAR.S|S|number1;[number2];…|Sample variance.
VARP|S|number1;[number2];…|Population variance.
VDB|F|cost;salvage;life;start_period;end_period;[factor];[no_switch]|Variable-declining balance depreciation.
VLOOKUP|R|lookup_value;table_array;col_index_num;[range_lookup]|Looks down the first column and returns a value from another column.
VSTACK|R|array1;[array2];…|Stacks arrays on top of each other.
WEEKDAY|D|serial_number;[return_type]|Day of the week as a number.
WEEKNUM|D|serial_number;[return_type]|Week number of the year.
WORKDAY|D|start_date;days;[holidays]|Date a number of working days away.
WORKDAY.INTL|D|start_date;days;[weekend];[holidays]|Working-day date with custom weekends.
WRAPCOLS|R|vector;wrap_count;[pad_with]|Wraps a vector into columns.
WRAPROWS|R|vector;wrap_count;[pad_with]|Wraps a vector into rows.
XIRR|F|values;dates;[guess]|Internal rate of return for irregular cash flows.
XLOOKUP|R|lookup_value;lookup_array;return_array;[if_not_found];[match_mode];[search_mode]|Looks up a value and returns the matching item.
XMATCH|R|lookup_value;lookup_array;[match_mode];[search_mode]|Position of a value in an array.
XNPV|F|rate;values;dates|Net present value for irregular cash flows.
XOR|L|logical1;[logical2];…|TRUE when an odd number of arguments are TRUE.
YEAR|D|serial_number|Year of a date.
YEARFRAC|D|start_date;end_date;[basis]|Fraction of the year between two dates.
`;

function guessCategory(name: string): string {
  if (/^(IM|BESSEL|BIN2|DEC2|HEX2|OCT2|BIT|ERF|COMPLEX|DELTA|GESTEP|CONVERT)/.test(name)) return CATEGORIES.E;
  if (/^D(AVERAGE|COUNT|GET|MAX|MIN|PRODUCT|STDEV|SUM|VAR)/.test(name)) return CATEGORIES.B;
  if (/^(COUP|ACCR|PRICE|YIELD|TBILL|DISC|INTRATE|RECEIVED|DURATION|MDURATION|DOLLARDE|DOLLARFR|FVSCHEDULE|PDURATION|RRI|ISPMT)/.test(name)) return CATEGORIES.F;
  if (/^IS|^ERROR\.TYPE|^INFO|^AREAS/.test(name)) return CATEGORIES.I;
  if (/DIST|INV|TEST|^CHI|^GAMMA|^BETA|^BINOM|^LOGNORM|^NEGBINOM|^POISSON|^EXPON|^WEIBULL|^F\.|^T\.|^Z|^CONFIDENCE|^COVAR|^PEARSON|^STEYX|^DEVSQ|^PROB|^LOGEST|^FISHER|^GAUSS|^PHI|^HYPGEOM|^CRITBINOM|^PERMUTATIONA|^PERCENTRANK|^QUARTILE|^SKEW|^VARA|^VARPA|^STDEVPA/.test(name)) return CATEGORIES.S;
  if (/B$|^ASC$|^DBCS$|^PHONETIC$|^ARRAYTOTEXT$/.test(name)) return CATEGORIES.T;
  return CATEGORIES.M;
}

function genericArgs(min: number, max: number): string[] {
  const out: string[] = [];
  const shown = Math.min(Math.max(min, 1), 4);
  const variadic = max >= 30;
  const n = variadic ? shown : max;
  for (let i = 1; i <= n; i++) out.push(i <= min ? `arg${i}` : `[arg${i}]`);
  if (variadic) out.push('…');
  return out;
}

let catalog: Map<string, FunctionInfo> | null = null;

export function functionCatalog(): Map<string, FunctionInfo> {
  if (catalog) return catalog;
  catalog = new Map();
  const described = new Map<string, FunctionInfo>();
  for (const line of TABLE.split('\n')) {
    if (!line.trim()) continue;
    const [name, cat, args, description] = line.split('|');
    described.set(name, { name, category: CATEGORIES[cat] ?? cat, args: args ? args.split(';') : [], description });
  }
  const names = new Set([...registry.keys(), 'LET', 'LAMBDA']);
  for (const name of [...names].sort()) {
    const known = described.get(name);
    if (known) {
      catalog.set(name, known);
      continue;
    }
    const def = registry.get(name);
    catalog.set(name, { name, category: guessCategory(name), args: def ? genericArgs(def.min, def.max) : [], description: '' });
  }
  return catalog;
}

export function functionInfo(name: string): FunctionInfo | undefined {
  return functionCatalog().get(name.toUpperCase());
}

/** Functions whose name starts with the prefix (case-insensitive). */
export function matchFunctions(prefix: string, limit = 60): FunctionInfo[] {
  const p = prefix.toUpperCase();
  const out: FunctionInfo[] = [];
  for (const [name, info] of functionCatalog()) {
    if (name.startsWith(p)) out.push(info);
    if (out.length >= limit) break;
  }
  return out;
}

export function functionCategories(): string[] {
  return [...new Set([...functionCatalog().values()].map((f) => f.category))].sort();
}

/** "SUM(number1, [number2], …)" */
export function signature(info: FunctionInfo): string {
  return `${info.name}(${info.args.join(', ')})`;
}

/** A short call example built from the parameter names (shown in the formula suggestions). */
export function exampleFor(info: FunctionInfo): string {
  let rangeCol = 0;
  let num = 0;
  const range = () => {
    const col = String.fromCharCode(65 + (rangeCol++ % 26));
    return `${col}2:${col}6`;
  };
  const sample = (raw: string): string | null => {
    const a = raw.replace(/^\[|\]$/g, '').toLowerCase();
    if (a === '…') return null;
    if (/criteria(?!_range)/.test(a)) return '">10"';
    if (/range|array|database|ref|values|known|data|matrix|vector|lookup_vector|result_vector/.test(a)) return range();
    if (/^value_if_true$/.test(a)) return '"Yes"';
    if (/^value_if_false|value_if_error|value_if_na|if_not_found$/.test(a)) return '"No"';
    if (/logical|condition|test/.test(a)) return 'A2>10';
    if (/find_text|old_text|delimiter|separator/.test(a)) return '"a"';
    if (/new_text|text\d*$|^text$|string/.test(a)) return '"text"';
    if (/start_date|^date$|serial_number|settlement/.test(a)) return 'A2';
    if (/end_date|maturity/.test(a)) return 'B2';
    if (/lookup_value|within_text/.test(a)) return 'A2';
    if (/col_index|row_index|index_num|column_num|row_num|^k$|nth/.test(a)) return '2';
    if (/range_lookup|exact|match_type/.test(a)) return '0';
    if (/digits|places|decimals|significance/.test(a)) return '2';
    if (/rate|probability|alpha/.test(a)) return '0.05';
    if (/nper|periods|months|days|years|count|num_chars|start_num/.test(a)) return '12';
    if (/pv|fv|pmt|principal|cost/.test(a)) return '1000';
    if (/number|value|x$|y$|base|power/.test(a)) return String([5, 10, 15, 20][num++ % 4]);
    return 'A2';
  };
  const required = info.args.filter((a) => !a.startsWith('[') && a !== '…');
  const args = required.length ? required : info.args.filter((a) => a !== '…').slice(0, 1);
  const variadic = info.args.includes('…');
  const parts = args.map(sample).filter((x): x is string => x !== null);
  if (variadic && parts.length === 1 && /^\d+$/.test(parts[0])) parts.push(String([5, 10, 15, 20][num++ % 4]), String([5, 10, 15, 20][num++ % 4]));
  return `${info.name}(${parts.join(', ')})`;
}
