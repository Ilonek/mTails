/* Value formatting and user-input parsing.
 *
 * Excel stores presentation separately from data: a cell keeps a raw value
 * and a number-format code such as "#,##0.00" or "yyyy-mm-dd".  This module
 * renders that pair into the string the grid shows, and parses what the user
 * types back into a typed value plus an inferred number format.
 *
 * The formatter covers the everyday subset of Excel codes: sections
 * (positive;negative;zero;text), colours, percentages, thousands grouping,
 * currency/locale brackets, scientific notation, fractions, dates, times and
 * elapsed time.  It is deliberately self-contained: no locale data files.
 */

const MS_DAY = 86400000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

export const ERROR_VALUES = new Set([
  "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CIRC!", "#SPILL!", "#CALC!",
]);

const BUILTIN = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
  9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
  18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
  22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

const COLORS = {
  black: "#000000", blue: "#0000ff", cyan: "#00ffff", green: "#00ff00",
  magenta: "#ff00ff", red: "#c00000", white: "#ffffff", yellow: "#ffff00",
};

const MONTHS = {
  en: {
    long: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    short: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  },
  pl: {
    long: ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września", "października", "listopada", "grudnia"],
    short: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru"],
  },
};

const WEEKDAYS = {
  en: {
    long: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    short: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  },
  pl: {
    long: ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"],
    short: ["niedz.", "pon.", "wt.", "śr.", "czw.", "pt.", "sob."],
  },
};

function language() {
  return document?.documentElement?.lang === "pl" ? "pl" : "en";
}

export function resolveFormat(code) {
  if (code === undefined || code === null || code === "") return "General";
  if (typeof code === "number") return BUILTIN[code] || "General";
  return String(code);
}

/* --------------------------------------------------------------- sections */
function splitSections(code) {
  const sections = [];
  let current = "";
  let quoted = false;
  let brackets = 0;
  for (let index = 0; index < code.length; index += 1) {
    const char = code[index];
    if (quoted) {
      current += char;
      if (char === '"') {
        if (code[index + 1] === '"') {
          current += '"';
          index += 1;
        } else {
          quoted = false;
        }
      }
      continue;
    }
    if (char === '"') {
      quoted = true;
      current += char;
      continue;
    }
    if (char === "[") {
      brackets += 1;
      current += char;
      continue;
    }
    if (char === "]") {
      brackets = Math.max(0, brackets - 1);
      current += char;
      continue;
    }
    if (char === ";" && brackets === 0) {
      sections.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  sections.push(current);
  return sections;
}

function sectionColor(section) {
  const match = /\[(black|blue|cyan|green|magenta|red|white|yellow)\]/i.exec(section || "");
  return match ? COLORS[match[1].toLowerCase()] : "";
}

const CONDITION_RE = /\[(<=|>=|<>|<|>|=)(-?\d+(?:\.\d+)?)\]/;

function conditionMatches(section, value) {
  const match = CONDITION_RE.exec(section);
  if (!match) return null;
  const number = Number(match[2]);
  switch (match[1]) {
    case "<": return value < number;
    case "<=": return value <= number;
    case ">": return value > number;
    case ">=": return value >= number;
    case "=": return value === number;
    case "<>": return value !== number;
    default: return null;
  }
}

function chooseSection(sections, value) {
  for (let index = 0; index < sections.length; index += 1) {
    if (conditionMatches(sections[index], value)) {
      return { section: sections[index], index, negative: value < 0, explicit: index > 0 };
    }
  }
  if (value < 0 && sections.length >= 2) {
    return { section: sections[1], index: 1, negative: true, explicit: true };
  }
  if (value === 0 && sections.length >= 3) {
    return { section: sections[2], index: 2, negative: false, explicit: true, zero: true };
  }
  return { section: sections[0], index: 0, negative: value < 0, explicit: false };
}

/* ------------------------------------------------------------- date tests */
function stripForDateTest(code) {
  let output = "";
  let quoted = false;
  for (let index = 0; index < code.length; index += 1) {
    const char = code[index];
    if (quoted) {
      if (char === '"') quoted = false;
      continue;
    }
    if (char === '"') { quoted = true; continue; }
    if (char === "\\") { index += 1; continue; }
    if (char === "_" || char === "*") { index += 1; continue; }
    if (char === "[") {
      const end = code.indexOf("]", index + 1);
      const content = end > index ? code.slice(index + 1, end) : "";
      if (/^[hms]+$/i.test(content)) output += content;
      index = end > index ? end : index;
      continue;
    }
    output += char;
  }
  return output;
}

export function isDateFormat(code) {
  const text = stripForDateTest(resolveFormat(code));
  if (!text) return false;
  if (/\[[hms]+\]/i.test(text)) return true;
  return /[ymdhs]/i.test(text);
}

export function isTimeFormat(code) {
  const text = stripForDateTest(resolveFormat(code));
  return /[hs]/i.test(text) || /\[[hms]+\]/i.test(text);
}

/* ---------------------------------------------------------- serial helpers */
export function serialToParts(serial) {
  const total = Math.round(Number(serial || 0) * MS_DAY);
  let dayIndex = Math.floor(total / MS_DAY);
  const timeMs = total - dayIndex * MS_DAY;
  if (dayIndex < 60) dayIndex += 1;
  const date = new Date(EXCEL_EPOCH + dayIndex * MS_DAY + timeMs);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekday: date.getUTCDay(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
    ms: date.getUTCMilliseconds(),
    serial: Number(serial || 0),
  };
}

export function serialFromParts(year, month, day, hours = 0, minutes = 0, seconds = 0) {
  const days = Math.floor((Date.UTC(year, month - 1, day) - EXCEL_EPOCH) / MS_DAY);
  let serial = days < 61 ? days - 1 : days;
  serial += (hours * 3600 + minutes * 60 + seconds) / 86400;
  return serial;
}

export function dateToSerial(date) {
  return serialFromParts(
    date.getFullYear(), date.getMonth() + 1, date.getDate(),
    date.getHours(), date.getMinutes(), date.getSeconds(),
  );
}

/* ------------------------------------------------------------ date format */
function tokenizeDate(section) {
  const tokens = [];
  const push = (type, value = "") => tokens.push({ type, value });
  let index = 0;
  while (index < section.length) {
    const char = section[index];
    if (char === '"') {
      let text = "";
      index += 1;
      while (index < section.length) {
        if (section[index] === '"') {
          if (section[index + 1] === '"') {
            text += '"';
            index += 2;
            continue;
          }
          index += 1;
          break;
        }
        text += section[index];
        index += 1;
      }
      push("literal", text);
      continue;
    }
    if (char === "\\") {
      if (section[index + 1] !== undefined) push("literal", section[index + 1]);
      index += 2;
      continue;
    }
    if (char === "_") {
      push("literal", " ");
      index += 2;
      continue;
    }
    if (char === "*") {
      index += 2;
      continue;
    }
    if (char === "[") {
      const end = section.indexOf("]", index + 1);
      const content = end > index ? section.slice(index + 1, end) : "";
      if (/^[hms]+$/i.test(content)) {
        push("elapsed", content[0].toLowerCase());
      }
      index = end > index ? end + 1 : index + 1;
      continue;
    }
    const slice = section.slice(index);
    if (/^AM\/PM/i.test(slice)) {
      push("ampm", "AM/PM");
      index += 5;
      continue;
    }
    if (/^A\/P/i.test(slice)) {
      push("ampm", "A/P");
      index += 3;
      continue;
    }
    if (/[ymdhs]/i.test(char)) {
      let run = 0;
      while (index + run < section.length && section[index + run].toLowerCase() === char.toLowerCase()) run += 1;
      push("field", char.toLowerCase().repeat(run));
      index += run;
      continue;
    }
    if (char === "." && /^\.0+/.test(slice)) {
      let run = 0;
      while (section[index + 1 + run] === "0") run += 1;
      push("fraction", String(run));
      index += run + 1;
      continue;
    }
    push("literal", char);
    index += 1;
  }
  return tokens;
}

function resolveMinutes(tokens) {
  const fields = tokens.filter((token) => token.type === "field" || token.type === "elapsed");
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.type !== "field" || token.value[0] !== "m") continue;
    const position = fields.indexOf(token);
    const previous = fields[position - 1];
    const next = fields[position + 1];
    const previousHour = previous && ((previous.type === "field" && previous.value[0] === "h")
      || (previous.type === "elapsed" && previous.value === "h"));
    const nextSecond = next && ((next.type === "field" && next.value[0] === "s")
      || (next.type === "elapsed" && next.value === "s"));
    token.type = previousHour || nextSecond ? "minute" : "month";
  }
}

export function formatDateSerial(serial, code) {
  const section = splitSections(resolveFormat(code))[0] || "General";
  const tokens = tokenizeDate(section);
  resolveMinutes(tokens);
  const parts = serialToParts(serial);
  const lang = language();
  const names = MONTHS[lang] || MONTHS.en;
  const days = WEEKDAYS[lang] || WEEKDAYS.en;
  const hasAmPm = tokens.some((token) => token.type === "ampm");
  let hour12 = parts.hours;
  if (hasAmPm) hour12 = parts.hours % 12 || 12;

  let output = "";
  for (const token of tokens) {
    if (token.type === "literal") {
      output += token.value;
      continue;
    }
    if (token.type === "ampm") {
      const pm = parts.hours >= 12;
      if (token.value === "A/P") output += pm ? "P" : "A";
      else output += pm ? "PM" : "AM";
      continue;
    }
    if (token.type === "minute") {
      const count = token.value.length;
      output += count > 1 ? String(parts.minutes).padStart(2, "0") : String(parts.minutes);
      continue;
    }
    if (token.type === "month") {
      const count = token.value.length;
      if (count >= 5) output += names.long[parts.month - 1][0];
      else if (count === 4) output += names.long[parts.month - 1];
      else if (count === 3) output += names.short[parts.month - 1];
      else if (count === 2) output += String(parts.month).padStart(2, "0");
      else output += String(parts.month);
      continue;
    }
    if (token.type === "fraction") {
      const count = Number(token.value) || 0;
      const value = String(Math.floor(parts.ms / 1000 * 10 ** count)).padStart(count, "0");
      output += value.slice(0, count);
      continue;
    }
    if (token.type === "elapsed") {
      const amount = token.value === "h"
        ? Math.floor(parts.serial * 24)
        : token.value === "m"
          ? Math.floor(parts.serial * 1440)
          : Math.floor(parts.serial * 86400);
      output += String(amount);
      continue;
    }
    const char = token.value[0];
    const count = token.value.length;
    switch (char) {
      case "y":
        output += count <= 2 ? String(parts.year).slice(-2) : String(parts.year).padStart(4, "0");
        break;
      case "d":
        if (count === 4) output += days.long[parts.weekday];
        else if (count === 3) output += days.short[parts.weekday];
        else if (count === 2) output += String(parts.day).padStart(2, "0");
        else output += String(parts.day);
        break;
      case "h": {
        const value = hasAmPm ? hour12 : parts.hours;
        output += count > 1 ? String(value).padStart(2, "0") : String(value);
        break;
      }
      case "s":
        output += count > 1 ? String(parts.seconds).padStart(2, "0") : String(parts.seconds);
        break;
      default:
        output += token.value;
    }
  }
  return output;
}

/* ---------------------------------------------------------- text sections */
function renderTextSection(section, value) {
  let output = "";
  let index = 0;
  let hasPlaceholder = false;
  while (index < section.length) {
    const char = section[index];
    if (char === '"') {
      index += 1;
      while (index < section.length) {
        if (section[index] === '"') {
          if (section[index + 1] === '"') {
            output += '"';
            index += 2;
            continue;
          }
          index += 1;
          break;
        }
        output += section[index];
        index += 1;
      }
      continue;
    }
    if (char === "\\") {
      if (section[index + 1] !== undefined) output += section[index + 1];
      index += 2;
      continue;
    }
    if (char === "_") { output += " "; index += 2; continue; }
    if (char === "*") { index += 2; continue; }
    if (char === "@") { output += String(value); hasPlaceholder = true; index += 1; continue; }
    if (char === "[") {
      const end = section.indexOf("]", index + 1);
      index = end > index ? end + 1 : index + 1;
      continue;
    }
    output += char;
    index += 1;
  }
  return hasPlaceholder ? output : String(value);
}

/* -------------------------------------------------------- numeric sections */
function findBracketEnd(text, start) {
  const end = text.indexOf("]", start + 1);
  return end < 0 ? text.length - 1 : end;
}

function numericTokens(section) {
  const tokens = [];
  const push = (type, value = "") => tokens.push({ type, value });
  let index = 0;
  while (index < section.length) {
    const char = section[index];
    if (char === '"') {
      let text = "";
      index += 1;
      while (index < section.length) {
        if (section[index] === '"') {
          if (section[index + 1] === '"') {
            text += '"';
            index += 2;
            continue;
          }
          index += 1;
          break;
        }
        text += section[index];
        index += 1;
      }
      push("literal", text);
      continue;
    }
    if (char === "\\") {
      if (section[index + 1] !== undefined) push("literal", section[index + 1]);
      index += 2;
      continue;
    }
    if (char === "_") { push("literal", " "); index += 2; continue; }
    if (char === "*") { index += 2; continue; }
    if (char === "[") {
      const end = findBracketEnd(section, index);
      const content = section.slice(index + 1, end);
      index = end + 1;
      if (content.startsWith("$")) {
        const symbol = content.replace(/^\$/, "").replace(/-[0-9A-Fa-f]+$/, "");
        if (symbol) push("literal", symbol);
      }
      continue;
    }
    if (char === "0" || char === "#" || char === "?") { push("digit", char); index += 1; continue; }
    if (char === ".") { push("dot", "."); index += 1; continue; }
    if (char === ",") { push("comma", ","); index += 1; continue; }
    if (char === "%") { push("percent", "%"); index += 1; continue; }
    if (char === "/") { push("slash", "/"); index += 1; continue; }
    if (char === "@") { push("at", "@"); index += 1; continue; }
    if ((char === "E" || char === "e") && (section[index + 1] === "+" || section[index + 1] === "-")) {
      push("exp", char.toUpperCase() + section[index + 1]);
      index += 2;
      continue;
    }
    push("literal", char);
    index += 1;
  }
  return tokens;
}

function tokensText(tokens) {
  let output = "";
  for (const token of tokens) {
    if (token.type === "literal") output += token.value;
    else if (token.type === "percent") output += "%";
    else if (token.type === "slash") output += "/";
    else if (token.type === "dot") output += ".";
    else if (token.type === "comma") output += ",";
    else if (token.type === "exp") output += token.value;
    else if (token.type === "at") output += "@";
    else if (token.type === "digit") output += token.value;
  }
  return output;
}

function groupDigits(text) {
  return text.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function roundTo(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function formatScientific(tokens, value, negative, autoMinus) {
  const expIndex = tokens.findIndex((token) => token.type === "exp");
  const firstDigit = tokens.findIndex((token) => token.type === "digit");
  const lastDigit = tokens.length - 1 - tokens.slice().reverse().findIndex((token) => token.type === "digit");
  if (expIndex < 0 || firstDigit < 0 || expIndex < firstDigit) return null;
  const mantissaTokens = tokens.slice(firstDigit, expIndex);
  const dotIndex = mantissaTokens.findIndex((token) => token.type === "dot");
  const decimalTokens = dotIndex >= 0
    ? mantissaTokens.slice(dotIndex + 1).filter((token) => token.type === "digit")
    : [];
  const decimals = decimalTokens.length;
  const required = decimalTokens.filter((token, index) => {
    if (token.value !== "0") return false;
    return decimalTokens.slice(0, index + 1).every((entry) => entry.value === "0");
  }).length;
  const expTokens = tokens.slice(expIndex + 1, lastDigit + 1).filter((token) => token.type === "digit");
  const prefix = tokens.slice(0, firstDigit).filter((token) => token.type !== "digit");
  const suffix = tokens.slice(lastDigit + 1);
  const expDigits = expTokens.length || 2;
  const signMode = tokens[expIndex].value;

  const number = Math.abs(value);
  let exponent = number === 0 ? 0 : Math.floor(Math.log10(number));
  let mantissa = number === 0 ? 0 : number / 10 ** exponent;
  mantissa = roundTo(mantissa, decimals);
  if (mantissa >= 10) {
    mantissa /= 10;
    exponent += 1;
  }
  let mantissaText = mantissa.toFixed(decimals);
  if (mantissaText.includes(".")) {
    const [wholePart, decimalPart = ""] = mantissaText.split(".");
    let trimmed = decimalPart;
    while (trimmed.length > required && trimmed.endsWith("0")) trimmed = trimmed.slice(0, -1);
    mantissaText = decimalPart.length ? wholePart + "." + trimmed : wholePart;
  }
  const exponentText = String(Math.abs(exponent)).padStart(expDigits, "0");
  const sign = exponent < 0 ? "-" : signMode === "E+" ? "+" : "";
  return (negative && autoMinus ? "-" : "") + tokensText(prefix)
    + mantissaText + "E" + sign + exponentText + tokensText(suffix);
}

function bestFraction(fraction, maxDenominator) {
  let bestNumerator = 0;
  let bestDenominator = 1;
  let bestError = Infinity;
  for (let denominator = 1; denominator <= maxDenominator; denominator += 1) {
    const numerator = Math.round(fraction * denominator);
    if (!numerator) continue;
    const error = Math.abs(fraction - numerator / denominator);
    if (error < bestError - 1e-12) {
      bestError = error;
      bestNumerator = numerator;
      bestDenominator = denominator;
    }
  }
  return { numerator: bestNumerator, denominator: bestDenominator };
}

function formatFraction(tokens, value, negative, autoMinus) {
  const slashIndex = tokens.findIndex((token) => token.type === "slash");
  if (slashIndex < 0) return null;
  const before = tokens.slice(0, slashIndex);
  const after = tokens.slice(slashIndex + 1);
  let numeratorEnd = before.length - 1;
  while (numeratorEnd >= 0 && before[numeratorEnd].type !== "digit") numeratorEnd -= 1;
  if (numeratorEnd < 0) return null;
  let numeratorStart = numeratorEnd;
  while (numeratorStart > 0 && before[numeratorStart - 1].type === "digit") numeratorStart -= 1;
  const numeratorTokens = before.slice(numeratorStart, numeratorEnd + 1);
  const wholeTokens = before.slice(0, numeratorStart);
  const hasWhole = wholeTokens.some((token) => token.type === "digit");

  let denominatorCount = 0;
  while (denominatorCount < after.length && after[denominatorCount].type === "digit") denominatorCount += 1;
  const denominatorTokens = after.slice(0, denominatorCount);
  let suffixStart = denominatorCount;
  let fixedDenominator = 0;
  if (!denominatorTokens.length) {
    const literal = after[0];
    if (literal && literal.type === "literal" && /^\d+$/.test(literal.value)) {
      fixedDenominator = Number(literal.value);
      suffixStart = 1;
    }
  }
  const suffix = after.slice(suffixStart);

  const number = Math.abs(value);
  const whole = Math.floor(number);
  const fraction = number - whole;
  let numerator;
  let denominator;
  if (fixedDenominator) {
    denominator = fixedDenominator;
    numerator = Math.round(fraction * denominator);
  } else if (hasWhole) {
    const maxDenominator = Math.max(1, 10 ** Math.max(1, denominatorTokens.length) - 1);
    const best = bestFraction(fraction, maxDenominator);
    numerator = best.numerator;
    denominator = best.denominator;
  } else {
    const maxDenominator = Math.max(1, 10 ** Math.max(1, denominatorTokens.length) - 1);
    const best = bestFraction(number, maxDenominator);
    numerator = best.numerator;
    denominator = best.denominator;
  }

  let wholeOut = whole;
  if (numerator >= denominator && denominator > 0) {
    if (hasWhole) {
      wholeOut += Math.floor(numerator / denominator);
      numerator %= denominator;
    }
  }

  let text = "";
  if (hasWhole) {
    if (wholeOut) text += String(wholeOut);
    else if (!numerator) text += "0";
  }
  if (numerator) {
    if (text) text += " ";
    const numeratorText = String(numerator).padStart(numeratorTokens.length, " ");
    const denominatorText = String(denominator).padStart(Math.max(1, denominatorTokens.length), " ");
    text += numeratorText + "/" + denominatorText;
  }
  if (!text) text = "0";
  return (negative && autoMinus ? "-" : "") + text + tokensText(suffix);
}

function formatNumericSection(section, value, negative, autoMinus) {
  const tokens = numericTokens(section);
  const digits = tokens.filter((token) => token.type === "digit");
  if (!digits.length) {
    return (negative && autoMinus ? "-" : "") + tokensText(tokens).replace(/@/g, String(value));
  }

  if (tokens.some((token) => token.type === "exp")) {
    const scientific = formatScientific(tokens, value, negative, autoMinus);
    if (scientific !== null) return scientific;
  }
  if (tokens.some((token) => token.type === "slash")) {
    const fraction = formatFraction(tokens, value, negative, autoMinus);
    if (fraction !== null) return fraction;
  }

  const firstDigit = tokens.findIndex((token) => token.type === "digit");
  const lastDigit = tokens.length - 1 - tokens.slice().reverse().findIndex((token) => token.type === "digit");
  const dotIndex = tokens.findIndex((token, index) => token.type === "dot" && index > firstDigit && index < lastDigit);
  const percentCount = tokens.filter((token) => token.type === "percent").length;

  const integerTokens = tokens.slice(firstDigit, dotIndex >= 0 ? dotIndex : lastDigit + 1);
  const decimalTokens = dotIndex >= 0 ? tokens.slice(dotIndex + 1, lastDigit + 1) : [];
  const prefixTokens = tokens.slice(0, firstDigit);
  const suffixTokens = tokens.slice(lastDigit + 1);

  let number = Math.abs(value) * 100 ** percentCount;
  const trailingCommas = integerTokens.filter((token, index) => {
    if (token.type !== "comma") return false;
    const rest = integerTokens.slice(index + 1);
    return !rest.some((entry) => entry.type === "digit");
  }).length;
  number /= 1000 ** trailingCommas;

  const integerDigits = integerTokens.filter((token) => token.type === "digit");
  const decimalDigits = decimalTokens.filter((token) => token.type === "digit");
  const grouping = integerTokens.some((token, index) => {
    if (token.type !== "comma") return false;
    const before = integerTokens.slice(0, index).some((entry) => entry.type === "digit");
    const after = integerTokens.slice(index + 1).some((entry) => entry.type === "digit");
    return before && after;
  });
  const minInteger = integerDigits.filter((token) => token.value === "0" || token.value === "?").length;
  const requiredDecimals = (() => {
    let count = 0;
    for (const token of decimalDigits) {
      if (token.value === "0") count += 1;
      else break;
    }
    return count;
  })();
  const decimals = decimalDigits.length;
  const rounded = roundTo(number, decimals);
  let [integerText, decimalText = ""] = rounded.toFixed(decimals).split(".");

  if (rounded === 0 && minInteger === 0) integerText = "";
  if (minInteger > integerText.length && integerText !== "") integerText = integerText.padStart(minInteger, "0");
  if (integerText && grouping) integerText = groupDigits(integerText);

  if (decimals) {
    decimalText = decimalText.padEnd(decimals, "0");
    const chars = decimalText.split("");
    for (let index = chars.length - 1; index >= requiredDecimals; index -= 1) {
      const placeholder = decimalDigits[index] || { value: "#" };
      if (chars[index] === "0" && placeholder.value === "#") chars[index] = "";
      else if (chars[index] === "0" && placeholder.value === "?") chars[index] = " ";
      else if (chars[index] === "0" && placeholder.value === "0" && index >= requiredDecimals) chars[index] = "0";
      else break;
    }
    decimalText = chars.join("");
    while (decimalText.endsWith(" ")) {
      // A '?' placeholder keeps its width; only trim empty '#' slots.
      break;
    }
  }

  let prefix = tokensText(prefixTokens).replace(/@/g, String(value));
  let suffix = tokensText(suffixTokens).replace(/@/g, String(value));
  let body = integerText;
  if (decimals > 0) body += "." + decimalText;
  if (!body && (prefix || suffix)) {
    // A literal-only zero section such as "-" still needs its text.
  }
  let output = (negative && autoMinus ? "-" : "") + prefix + body + suffix;
  if (!integerText && !decimalText && !prefix && !suffix) output = "";
  return output;
}

/* ------------------------------------------------------------- general fmt */
export function generalFormat(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  const number = Number(value);
  if (!Number.isFinite(number)) return "#NUM!";
  if (number === 0) return "0";
  if (Number.isInteger(number) && Math.abs(number) < 1e11) return String(number);
  const abs = Math.abs(number);
  if (abs >= 1e11 || abs < 1e-9) {
    return number.toExponential(5)
      .replace(/e([+-])(\d)$/, "E$10$2")
      .replace(/e([+-])/i, "E$1")
      .replace(/\.?0+E/, "E")
      .replace("e", "E");
  }
  let text = Number(number.toPrecision(11)).toString();
  if (text.includes("e")) {
    text = number.toFixed(20).replace(/0+$/, "").replace(/\.$/, "");
  }
  return text;
}

/* --------------------------------------------------------- public format */
export function formatNumberParts(value, code) {
  const format = resolveFormat(code);
  if (typeof value === "string") {
    const sections = splitSections(format);
    const section = sections.length >= 4 ? sections[3] : format === "General" ? "@" : sections[0];
    return { text: format === "General" ? value : renderTextSection(section, value), color: sectionColor(section) };
  }
  if (typeof value === "boolean") {
    return { text: value ? "TRUE" : "FALSE", color: "" };
  }
  const number = Number(value);
  if (!Number.isFinite(number)) return { text: "#NUM!", color: "", error: true };
  if (!format || format === "General" || format === "@") return { text: generalFormat(number), color: "" };
  if (isDateFormat(format)) {
    const sections = splitSections(format);
    const choice = chooseSection(sections, number);
    return { text: formatDateSerial(number, choice.section), color: sectionColor(choice.section) };
  }
  const sections = splitSections(format);
  const choice = chooseSection(sections, number);
  const color = sectionColor(choice.section);
  const text = formatNumericSection(choice.section, number, choice.negative, !choice.explicit);
  return { text, color };
}

export function formatValue(value, code) {
  if (value === null || value === undefined) return "";
  return formatNumberParts(value, code).text;
}

export function formatCell(cell, options = {}) {
  if (!cell) return { text: "", kind: "empty", color: "", align: "left" };
  if (cell.v === undefined || cell.v === null || cell.v === "") {
    if (cell.f) return { text: "", kind: "number", color: "", align: "right" };
    return { text: "", kind: "empty", color: "", align: "left" };
  }
  const code = resolveFormat(cell.s?.numFmt || options.numFmt || "General");
  const value = cell.v;
  if (typeof value === "boolean") {
    return { text: value ? "TRUE" : "FALSE", kind: "boolean", color: "", align: "center" };
  }
  if (typeof value === "number") {
    const parts = formatNumberParts(value, code);
    return {
      text: parts.text,
      kind: parts.error ? "error" : isDateFormat(code) ? "date" : "number",
      color: parts.color,
      align: "right",
    };
  }
  if (typeof value === "string" && ERROR_VALUES.has(value)) {
    return { text: value, kind: "error", color: "#c62828", align: "center" };
  }
  const parts = formatNumberParts(value, code);
  return { text: parts.text, kind: "text", color: parts.color, align: "left" };
}

/* ---------------------------------------------------------- input parsing */
const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

function parseNumberInput(text) {
  let value = text.trim();
  if (!value) return null;
  if (/^[+-]?0\d/.test(value)) return null;
  let negative = false;
  if (/^\(.*\)$/.test(value)) {
    negative = true;
    value = value.slice(1, -1).trim();
  }
  let percent = 0;
  while (value.endsWith("%")) {
    percent += 1;
    value = value.slice(0, -1).trim();
  }
  let currency = "";
  const currencyMatch = /^([$€£¥₽]|zł|PLN|USD|EUR)\s*/i.exec(value);
  if (currencyMatch) {
    currency = currencyMatch[1];
    value = value.slice(currencyMatch[0].length);
  }
  const trailingCurrency = /([$€£¥₽]|zł|PLN|USD|EUR)$/i.exec(value);
  if (!currency && trailingCurrency) {
    currency = trailingCurrency[1];
    value = value.slice(0, -trailingCurrency[0].length).trim();
  }
  let grouping = false;
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(value)) {
    grouping = true;
    value = value.replace(/,/g, "");
  } else if (!value.includes(".") && /^[+-]?\d+,\d{1,2}$/.test(value)) {
    value = value.replace(",", ".");
  } else if (/^[+-]?\d{1,3}(\s\d{3})+(\.\d+)?$/.test(value)) {
    grouping = true;
    value = value.replace(/\s/g, "");
  } else if (value.includes(",")) {
    return null;
  }
  if (!NUMBER_RE.test(value)) return null;
  let number = Number(value);
  if (!Number.isFinite(number)) return null;
  if (negative) number = -number;
  const decimals = (value.split(".")[1] || "").length;
  let numFmt;
  if (percent) {
    number /= 100 ** percent;
    numFmt = decimals ? "0.00%" : "0%";
  } else if (currency) {
    const symbol = currency.length > 1 ? `${currency} ` : currency;
    numFmt = `${symbol}#,##0.00`;
  } else if (grouping) {
    numFmt = "#,##0" + (decimals ? "." + "0".repeat(decimals) : "");
  }
  return { kind: "number", value: number, numFmt };
}

function validDate(year, month, day) {
  if (!year || month < 1 || month > 12 || day < 1) return false;
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function parseDateInput(text) {
  const value = text.trim();
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(value);
  if (match) {
    const [, year, month, day, hours = "0", minutes = "0", seconds = "0"] = match;
    if (!validDate(+year, +month, +day)) return null;
    const serial = serialFromParts(+year, +month, +day, +hours, +minutes, +seconds);
    return {
      kind: "number",
      value: serial,
      numFmt: hours !== undefined ? (seconds !== undefined ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd hh:mm") : "yyyy-mm-dd",
    };
  }

  match = /^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(value);
  if (match) {
    let [, first, second, year, hours = "0", minutes = "0", seconds = "0"] = match;
    let day = +first;
    let month = +second;
    const pl = language() === "pl";
    if (+first > 12) {
      day = +first;
      month = +second;
    } else if (+second > 12) {
      month = +first;
      day = +second;
    } else if (pl) {
      day = +first;
      month = +second;
    } else {
      month = +first;
      day = +second;
    }
    if (+year < 100) year = String(2000 + Number(year));
    if (!validDate(+year, month, day)) return null;
    const serial = serialFromParts(+year, month, day, +hours, +minutes, +seconds);
    const datePart = pl ? "dd.mm.yyyy" : "m/d/yyyy";
    return {
      kind: "number",
      value: serial,
      numFmt: hours !== undefined ? `${datePart} hh:mm` : datePart,
    };
  }

  match = /^(\d{1,2})-([A-Za-z]{3,})-(\d{2,4})$/.exec(value);
  if (match) {
    const [, day, monthName, yearText] = match;
    const year = yearText.length <= 2 ? 2000 + Number(yearText) : Number(yearText);
    const lang = language();
    const names = MONTHS[lang] || MONTHS.en;
    const lower = monthName.toLowerCase();
    let month = names.long.findIndex((name) => name.toLowerCase().startsWith(lower.slice(0, 3))) + 1;
    if (!month) month = names.short.findIndex((name) => name.toLowerCase() === lower) + 1;
    if (!month || !validDate(year, month, +day)) return null;
    return { kind: "number", value: serialFromParts(year, month, +day), numFmt: "d-mmm-yy" };
  }

  match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(value);
  if (match) {
    let [, hours, minutes, seconds = "0", suffix] = match;
    let hour = +hours;
    if (suffix && /pm/i.test(suffix) && hour < 12) hour += 12;
    if (suffix && /am/i.test(suffix) && hour === 12) hour = 0;
    if (hour > 23 || +minutes > 59 || +seconds > 59) return null;
    return {
      kind: "number",
      value: (hour * 3600 + +minutes * 60 + +seconds) / 86400,
      numFmt: seconds !== "0" ? "hh:mm:ss" : "hh:mm",
    };
  }
  return null;
}

export function parseUserInput(text) {
  const raw = String(text ?? "");
  if (!raw.length) return { kind: "empty" };
  if (raw.startsWith("=")) return { kind: "formula", formula: raw };
  if (raw.startsWith("'")) return { kind: "text", value: raw.slice(1) };
  if (/^true$/i.test(raw)) return { kind: "boolean", value: true };
  if (/^false$/i.test(raw)) return { kind: "boolean", value: false };
  const date = parseDateInput(raw);
  if (date) return date;
  const number = parseNumberInput(raw);
  if (number) return number;
  return { kind: "text", value: raw };
}

/* --------------------------------------------------------------- metadata */
export const STANDARD_FORMATS = [
  { label: "General", code: "General" },
  { label: "Number", code: "#,##0.00" },
  { label: "Currency", code: "$#,##0.00" },
  { label: "Accounting", code: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' },
  { label: "Short Date", code: "m/d/yyyy" },
  { label: "Long Date", code: "mmmm d, yyyy" },
  { label: "Time", code: "h:mm:ss AM/PM" },
  { label: "Percentage", code: "0.00%" },
  { label: "Fraction", code: "# ?/?" },
  { label: "Scientific", code: "0.00E+00" },
  { label: "Text", code: "@" },
];

export function formatCategory(code) {
  const text = resolveFormat(code);
  if (text === "General") return "general";
  if (isDateFormat(text)) return isTimeFormat(text) && !/[yd]/i.test(stripForDateTest(text)) ? "time" : "date";
  if (/%/.test(text)) return "percentage";
  if (/[Ee][+-]/.test(text)) return "scientific";
  if (/\?\/\?|\/\d/.test(text)) return "fraction";
  if (/@/.test(text)) return "text";
  if (/[$€£¥]/.test(text) || /\[\$/.test(text)) return "currency";
  if (/,/.test(text)) return "accounting";
  return "number";
}

export function formatPreview(code, value = 1234.567) {
  return formatValue(value, resolveFormat(code));
}
