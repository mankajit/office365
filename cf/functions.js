/* global CustomFunctions */
(function loadLib(){
  if (typeof window.DateConverter === 'undefined') {
    var s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/@remotemerge/nepali-date-converter@1/dist/ndc-browser.js';
    document.head.appendChild(s);
  }
})();

function waitForLib(cb){
  if (typeof window.DateConverter !== 'undefined') return cb();
  var tries = 0;
  var t = setInterval(function(){
    if (typeof window.DateConverter !== 'undefined'){ clearInterval(t); cb(); }
    if (++tries > 200) { clearInterval(t); throw new Error('Nepali date library failed to load.'); }
  }, 25);
}

function toISOFromExcelSerial(n){
  // Excel serial days since 1899-12-30
  var epoch = new Date(Date.UTC(1899,11,30));
  var d = new Date(epoch.getTime() + Math.round(n) * 86400000);
  return d.toISOString().slice(0,10);
}
function parseAsAD(dateOrYear, month, day){
  if (month == null && day == null){
    if (typeof dateOrYear === 'number') return toISOFromExcelSerial(dateOrYear);
    if (typeof dateOrYear === 'string') return dateOrYear.slice(0,10);
    if (dateOrYear && dateOrYear.getFullYear) return dateOrYear.toISOString().slice(0,10);
    throw new Error('Invalid input');
  } else {
    var y = Number(dateOrYear), m = Number(month), d = Number(day);
    var mm = String(m).padStart(2,'0'); var dd = String(d).padStart(2,'0');
    return [y,mm,dd].join('-');
  }
}

/**
 * Convert AD to BS (YYYY-MM-DD)
 * @customfunction
 * @param {any} dateOrYear Excel date/ISO text OR Year
 * @param {number} [month] Optional month
 * @param {number} [day] Optional day
 * @returns {string}
 */
function ADTOBS(dateOrYear, month, day){
  return new CustomFunctions.AsyncResult(async (resolve, reject)=>{
    try{
      await new Promise(waitForLib);
      var iso = parseAsAD(dateOrYear, month, day);
      var res = new window.DateConverter(iso).toBs();
      var out = [res.year, String(res.month).padStart(2,'0'), String(res.date).padStart(2,'0')].join('-');
      resolve(out);
    }catch(e){ reject(e); }
  });
}

/**
 * Convert AD to BS (long format)
 * @customfunction
 * @param {any} dateOrYear
 * @param {number} [month]
 * @param {number} [day]
 * @returns {string}
 */
function ADTOBS_LONG(dateOrYear, month, day){
  return new CustomFunctions.AsyncResult(async (resolve, reject)=>{
    try{
      await new Promise(waitForLib);
      var iso = parseAsAD(dateOrYear, month, day);
      var res = new window.DateConverter(iso).toBs();
      var months = ['Baisakh','Jestha','Ashadh','Shrawan','Bhadra','Ashwin','Kartik','Mangsir','Poush','Magh','Falgun','Chaitra'];
      var out = String(res.date).padStart(2,'0') + ' ' + months[res.month-1] + ' ' + res.year;
      resolve(out);
    }catch(e){ reject(e); }
  });
}

/**
 * Convert BS text (YYYY-MM-DD) to AD serial
 * @customfunction
 * @param {string} bsText
 * @returns {number}
 */
function BSTOAD(bsText){
  return new CustomFunctions.AsyncResult(async (resolve, reject)=>{
    try{
      await new Promise(waitForLib);
      var res = new window.DateConverter(bsText).toAd();
      var d = new Date(Date.UTC(res.year, res.month-1, res.date));
      var epoch = new Date(Date.UTC(1899,11,30));
      var serial = Math.round((d - epoch) / 86400000);
      resolve(serial);
    }catch(e){ reject(e); }
  });
}

CustomFunctions.associate('ADTOBS', ADTOBS);
CustomFunctions.associate('ADTOBS_LONG', ADTOBS_LONG);
CustomFunctions.associate('BSTOAD', BSTOAD);
