const fs = require('node:fs');
const assert = require('node:assert/strict');
const path = require('node:path');
const { jsPDF } = require('../app/static/vendor/jspdf.umd.min.js');
const forecast = require('../app/static/roi-forecast.js');
const create = require('../app/static/roi-pdf.js');
const inputs = {leads:5,conversion:5,value:62000,rate:5490,mode:'project',churn:10,churnEnabled:true,currency:'GBP'};
const labels = Array.from({length:60},(_,i)=>new Date(Date.UTC(2026,9+i,1)).toLocaleDateString('en-GB',{month:'short',year:'numeric',timeZone:'UTC'}));
const logo = 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname,'../app/static/img/logo-report.png')).toString('base64');
for (const mode of ['project','retainer']) {
 const data={...inputs,mode};
 const pdf=create({forecast:forecast(data),inputs:data,labels,logo},jsPDF);
 assert.equal(pdf.getNumberOfPages(),3);
 const output=Buffer.from(pdf.output('arraybuffer'));
 assert.equal(output.subarray(0,5).toString(),'%PDF-');
 if (process.argv[2]) fs.writeFileSync(process.argv[2].replace('.pdf',`-${mode}.pdf`),output);
}
console.log('PDF generation passed for projects and retainers: three pages with embedded logo.');
