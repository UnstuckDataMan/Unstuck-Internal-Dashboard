const fs=require('node:fs');
const path=require('node:path');
const assert=require('node:assert/strict');
const {jsPDF}=require('../app/static/vendor/jspdf.umd.min.js');
const forecast=require('../app/static/roi-forecast.js');
const create=require('../app/static/roi-pdf.js');
const logo='data:image/png;base64,'+fs.readFileSync(path.join(__dirname,'../app/static/img/logo-report.png')).toString('base64');
const out=process.argv[2];if(out)fs.mkdirSync(out,{recursive:true});
let count=0;
for(const currency of ['GBP','USD'])for(const mode of ['project','retainer'])for(const churn of [0,30,100])for(const large of [false,true])for(const horizon of [1,3,5]) {
 const inputs={currency,mode,churn,horizon,churnEnabled:churn!==0,leads:large?1000:3,conversion:large?100:5,value:large?1000000:25000,rate:currency==='USD'?7990:5490};
 const labels=Array.from({length:60},(_,i)=>new Date(Date.UTC(2026,11+i,1)).toLocaleDateString('en-GB',{month:'short',year:'numeric',timeZone:'UTC'}));
 const f=forecast(inputs),pdf=create({forecast:f,inputs,labels,logo},jsPDF);
 assert.equal(pdf.getNumberOfPages(),3);
 const filename=`${currency}-${mode}-${churn}-${large?'large':'normal'}-${horizon}yr.pdf`;
 if(out){fs.writeFileSync(path.join(out,filename),Buffer.from(pdf.output('arraybuffer')));fs.writeFileSync(path.join(out,filename+'.json'),JSON.stringify({inputs,forecast:f}));}
 count++;
}
console.log(`${count} PDF stress exports passed: 1/3/5-year headlines, both currencies, both fee types, churn 0/30/100%, large amounts and December year rollover.`);
