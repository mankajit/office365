// Self-contained AD <-> BS converter (demo range).
// IMPORTANT: For wide-range accuracy, extend bsMonthData for BS 1970–2100 with authoritative data.

const ADBS_internal = (()=>{
  // Minimal month-length data for demonstration (BS 2075–2085).
  // Replace/extend with full table for production use.
  const bsMonthData = {
    2075: [31,31,32,31,31,30,29,30,29,30,29,31],
    2076: [31,31,32,31,31,30,29,30,29,30,29,31],
    2077: [31,32,31,32,31,30,29,30,29,30,29,31],
    2078: [31,31,32,31,31,30,29,30,29,30,29,31],
    2079: [31,31,32,31,31,30,29,30,29,30,29,31],
    2080: [31,31,32,31,31,30,29,30,29,30,29,31],
    2081: [31,32,31,32,31,30,29,30,29,30,29,31],
    2082: [31,31,32,31,31,30,29,30,29,30,29,31],
    2083: [31,31,32,31,31,30,29,30,29,30,29,31],
    2084: [31,32,31,32,31,30,29,30,29,30,29,31],
    2085: [31,31,32,31,31,30,29,30,29,30,29,31]
  };

  // Anchor mapping: 2023-01-01 AD == 2079-09-17 BS (illustrative; adjust if needed).
  const anchorAD = { y:2023, m:1, d:1 };
  const anchorBS = { y:2079, m:9, d:17 };

  function toJulianDay(y, m, d){
    const a = Math.floor((14 - m) / 12);
    const y2 = y + 4800 - a;
    const m2 = m + 12 * a - 3;
    return d + Math.floor((153 * m2 + 2) / 5) + 365 * y2 + Math.floor(y2 / 4) - Math.floor(y2 / 100) + Math.floor(y2 / 400) - 32045;
  }

  function adToBs(ad){
    const jdAd = toJulianDay(ad.year, ad.month, ad.day);
    const jdAnchor = toJulianDay(anchorAD.y, anchorAD.m, anchorAD.d);
    let diff = jdAd - jdAnchor;

    let y = anchorBS.y, m = anchorBS.m, d = anchorBS.d;
    if(!bsMonthData[y]) throw new Error('BS data missing for year '+y);

    while(diff !== 0){
      if(diff > 0){
        d++;
        const ml = bsMonthData[y] && bsMonthData[y][m-1];
        if(!ml) throw new Error('BS data missing for year '+y);
        if(d > ml){
          d = 1; m++;
          if(m > 12){ m = 1; y++; if(!bsMonthData[y]) throw new Error('BS data missing for year '+y); }
        }
        diff--;
      } else {
        d--;
        if(d < 1){
          m--; if(m < 1){ y--; m = 12; if(!bsMonthData[y]) throw new Error('BS data missing for year '+y); }
          d = bsMonthData[y][m-1];
        }
        diff++;
      }
    }
    return { year:y, month:m, day:d };
  }

  function tryParseExcelCell(cell){
    if(cell == null) return null;
    if(cell instanceof Date){
      return { year: cell.getFullYear(), month: cell.getMonth()+1, day: cell.getDate() };
    }
    // Excel serial date number (approx, days since 1899-12-30)
    if(typeof cell === 'number'){
      const base = new Date(Date.UTC(1899,11,30));
      const dt = new Date(base.getTime() + Math.round(cell)*86400000);
      return { year: dt.getUTCFullYear(), month: dt.getUTCMonth()+1, day: dt.getUTCDate() };
    }
    // ISO string YYYY-MM-DD
    if(typeof cell === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(cell)){
      const [y,m,d] = cell.split('-').map(Number);
      return { year:y, month:m, day:d };
    }
    const parsed = new Date(cell);
    if(!isNaN(parsed)) return { year: parsed.getUTCFullYear(), month: parsed.getUTCMonth()+1, day: parsed.getUTCDate() };
    return null;
  }

  function formatBsISO(bs){
    return `${bs.year}-${String(bs.month).padStart(2,'0')}-${String(bs.day).padStart(2,'0')}`;
  }
  function formatBsLong(bs){
    const months = ['Baisakh','Jestha','Ashadh','Shrawan','Bhadra','Ashwin','Kartik','Mangsir','Poush','Magh','Falgun','Chaitra'];
    return `${String(bs.day).padStart(2,'0')} ${months[bs.month-1]} ${bs.year}`;
  }

  return { adToBs, tryParseExcelCell, formatBsISO, formatBsLong };
})();
