import '../converter.js';

/**
 * Convert an AD date to BS. Accepts either a single Excel date/cell value OR (year,month,day).
 * @customfunction
 * @param {*} dateOrYear Excel date/cell (serial, Date, or ISO text) OR Year number
 * @param {number} [month] Optional month (1-12) if first arg is year
 * @param {number} [day] Optional day (1-31) if first arg is year
 * @returns {string} BS date as YYYY-MM-DD
 */
function ADTOBS(dateOrYear, month, day) {
  let ad;
  if (month == null && day == null) {
    // single-arg mode
    ad = ADBS_internal.tryParseExcelCell(dateOrYear);
    if(!ad) throw new Error('Could not parse date from cell/value.');
  } else {
    const y = Number(dateOrYear);
    const m = Number(month);
    const d = Number(day);
    ad = { year: y, month: m, day: d };
  }
  const bs = ADBS_internal.adToBs(ad);
  return ADBS_internal.formatBsISO(bs);
}

// Optional helper that returns long form
/**
 * @customfunction
 * @param {*} dateOrYear
 * @param {number} [month]
 * @param {number} [day]
 * @returns {string}
 */
function ADTOBS_LONG(dateOrYear, month, day) {
  let ad;
  if (month == null && day == null) {
    ad = ADBS_internal.tryParseExcelCell(dateOrYear);
    if(!ad) throw new Error('Could not parse date from cell/value.');
  } else {
    ad = { year: Number(dateOrYear), month: Number(month), day: Number(day) };
  }
  const bs = ADBS_internal.adToBs(ad);
  return ADBS_internal.formatBsLong(bs);
}

if (typeof module !== 'undefined') module.exports = { ADTOBS, ADTOBS_LONG };
