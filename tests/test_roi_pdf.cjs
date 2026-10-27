const fs = require('node:fs');
const assert = require('node:assert/strict');
const path = require('node:path');
const { jsPDF } = require('../app/static/vendor/jspdf.umd.min.js');
const forecast = require('../app/static/roi-forecast.js');
const create = require('../app/static/roi-pdf.js');
const inputs = {leads:5,conversion:5,value:62000,rate:5490,mode:'project',churn:10,churnEnabled:true,currency:'GBP'};
const labels = Array.from({length:60},(_,i)=>new Date(Date.UTC(2026,9+i,1)).toLocaleDateString('en-GB',{month:'short',year:'numeric',timeZone:'UTC'}));
const logo = 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname,'../app/static/img/logo-report.png')).toString('base64');
for (const mode of ['project','retainer']) for (const horizon of [1,3,5]) {
 const data={...inputs,mode,horizon};
 const text=[];
 function RecordingPdf(options) {
  const doc=new jsPDF(options),write=doc.text.bind(doc);
  doc.text=(value,x,y,...rest)=>{text.push({value,x,y,page:doc.getCurrentPageInfo().pageNumber});return write(value,x,y,...rest);};
  return doc;
 }
 const result=forecast(data);
 const pdf=create({forecast:result,inputs:data,labels,logo},RecordingPdf);
 assert.equal(pdf.getNumberOfPages(),3);
 const output=Buffer.from(pdf.output('arraybuffer'));
 assert.equal(output.subarray(0,5).toString(),'%PDF-');
 const at=(page,x,y)=>text.find(item=>item.page===page&&item.x===x&&item.y===y)?.value;
 const selected=result.years[horizon-1],money=new Intl.NumberFormat('en-GB',{style:'currency',currency:'GBP'});
 const period=`${horizon} ${horizon===1?'year':'years'}`;
 assert.equal(at(1,18,71),`${labels[0]} - ${labels[horizon*12-1]}  |  GBP  |  Over ${period}`);
 assert.equal(at(1,26,132),selected.ratio.toFixed(2)+' : 1');
 assert.equal(at(1,188,155),money.format(selected.accumulated));
 assert.equal(at(1,188,164),money.format(selected.investment));
 assert.equal(at(1,188,173),money.format(selected.accumulated-selected.investment));
 assert.equal(at(1,188,182),String(horizon*3));
 assert.equal(at(3,18,70),`Headline forecast period: ${period}`);
 assert.equal(at(2,18,164),'First 12 months - monthly breakdown');
 if (process.argv[2]) fs.writeFileSync(process.argv[2].replace('.pdf',`-${mode}-${horizon}yr.pdf`),output);
}
assert.throws(()=>create({forecast:forecast(inputs),inputs:{...inputs,horizon:2},labels,logo},jsPDF),RangeError);
console.log('PDF generation passed: projects/retainers across 1/3/5 years, aligned headline revenue/fees/ratio/clients, clear period and preserved monthly breakdown.');
