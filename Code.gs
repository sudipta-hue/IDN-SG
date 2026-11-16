/** IDN HITES live dashboard read-only API.
 * Paste this into Extensions > Apps Script in the linked spreadsheet.
 * Deploy as Web app: Execute as Me; Who has access: Anyone.
 */
function doGet() {
  const ss = SpreadsheetApp.openById('149bSf_wDVsF0H1fRprVPMUvQdp5Z6NjjO9DE1ehxJVU');
  const out = {};
  const specs = {
    ongoing: {name:'ONGOING', start:4, cols:28, header:false},
    tender: {name:'Tendering Stage', start:2, cols:12, header:false},
    completed: {name:'COMPLETED PHYSICALLY', start:3, cols:12, header:false}
  };
  Object.keys(specs).forEach(key => {
    const s = specs[key], sh = ss.getSheetByName(s.name);
    if (!sh) { out[key] = []; return; }
    const values = sh.getRange(s.start, 1, Math.max(1, sh.getLastRow()-s.start+1), s.cols).getDisplayValues();
    out[key] = values;
  });
  return ContentService.createTextOutput(JSON.stringify({
    updatedAt: new Date().toISOString(),
    ongoing: out.ongoing.filter(r => r[1] && r[0] && !isNaN(Number(r[0]))).map(r => ({
      sno:r[0],name:r[1],contractor:r[2],dprCost:r[3],loa:r[4],location:r[5],client:r[6],
      start:r[7],scheduled:r[9],expected:r[10],delay:r[11],planned:r[12],physical:r[13],
      monthlyRequired:r[14],certified:r[15],financial:r[16],balanceCost:r[17],fundBalance:r[18],
      team:r[19],drawing:r[20],quality:r[21],labour:r[22],other:r[23],action:r[24],
      equipment:r[25],hindrances:r[26],mitigation:r[27]
    })),
    tender: out.tender.filter(r => r[1] && String(r[1]).toLowerCase() !== 'total cost').map(r => ({
      sno:r[0],name:r[1],cost:r[2],prep:r[3],approval:r[4],nit:r[5],prebid:r[6],
      opening:r[7],technical:r[8],financialEval:r[9],negotiation:r[10],finalisation:r[11]
    })),
    completed: out.completed.filter(r => r[1] && r[0]).map(r => ({
      sno:r[0],name:r[1],loa:r[2],location:r[3],client:r[4],start:r[5],completion:r[6],
      contractor:r[7],eot:r[8],deviations:r[9],finalBill:r[10],fundBalance:r[11]
    }))
  })).setMimeType(ContentService.MimeType.JSON);
}
