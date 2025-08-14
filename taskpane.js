(async function(){
  document.getElementById('convertBtn').addEventListener('click', onConvertClicked);
  document.getElementById('insertBtn').addEventListener('click', onInsertClicked);
  document.getElementById('convertRangeBtn').addEventListener('click', onConvertRangeClicked);

  async function onConvertClicked(){
    const val = document.getElementById('adDate').value;
    if(!val){ alert('Choose a date first'); return; }
    const [y,m,d] = val.split('-').map(Number);
    try{
      const bs = ADBS_internal.adToBs({year:y,month:m,day:d});
      document.getElementById('result').innerText = ADBS_internal.formatBsLong(bs);
    }catch(e){
      document.getElementById('result').innerText = 'Error: '+e.message;
    }
  }

  async function onInsertClicked(){
    const val = document.getElementById('adDate').value;
    if(!val){ alert('Choose a date first'); return; }
    const [y,m,d] = val.split('-').map(Number);
    const bs = ADBS_internal.adToBs({year:y,month:m,day:d});
    const text = ADBS_internal.formatBsISO(bs);
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[text]];
      await context.sync();
    });
  }

  async function onConvertRangeClicked(){
    await Excel.run(async (context)=>{
      const range = context.workbook.getSelectedRange();
      range.load('values, rowCount, columnCount');
      await context.sync();

      const out = [];
      for(let r=0;r<range.rowCount;r++){
        out[r] = [];
        for(let c=0;c<range.columnCount;c++){
          const cell = range.values[r][c];
          let dateObj = ADBS_internal.tryParseExcelCell(cell);
          if(!dateObj){ out[r][c] = ''; continue; }
          try{
            const bs = ADBS_internal.adToBs(dateObj);
            out[r][c] = ADBS_internal.formatBsISO(bs);
          }catch(e){ out[r][c] = 'ERR'; }
        }
      }
      range.values = out;
      await context.sync();
      document.getElementById('rangeResult').innerText = 'Converted '+range.rowCount+'×'+range.columnCount+' cells.';
    });
  }
})();