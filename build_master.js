const ExcelJS = require('exceljs');
const fs = require('fs');
const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));

(async function() {
    const wb = new ExcelJS.Workbook();
    const h = { font: { bold: true, color: { argb: 'FFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1565C0' } }, alignment: { horizontal: 'center' } };
    const h2 = { font: { bold: true, color: { argb: 'FFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } }, alignment: { horizontal: 'center' } };

    // Compute stats
    let withEmail=0, withPhone=0, work=0, broken=0, noWeb=0;
    const cats = {};
    for (const b of businesses) {
        if(b.email&&b.email.length>0) withEmail++;
        if((b.phone&&b.phone.length>0)||(b.mobile&&b.mobile.length>0)) withPhone++;
        const w = b.verifiedWebsite||b.website||'';
        if(!w) noWeb++; else if(b.websiteWorking===true) work++; else broken++;
        cats[b.category||'Uncategorized'] = (cats[b.category||'Uncategorized']||0)+1;
    }
    const sc = Object.entries(cats).sort((a,b)=>b[1]-a[1]);

    // 1. Summary
    const ws = wb.addWorksheet('Summary');
    ws.mergeCells('A1:B1');
    ws.getCell('A1').value = 'KIGALI BUSINESS DIRECTORY — MASTER FILE';
    ws.getCell('A1').style = { font: { bold: true, size: 16, color: { argb: '0D47A1' } } };
    const data = [
        ['Total Businesses', businesses.length], ['With Website (Verified)', work+broken],
        ['Working Websites', work], ['No Website', noWeb], ['With Email', withEmail],
        ['With Phone', withPhone], ['Total Categories', sc.length], ['', ''],
        ['Website Coverage', Math.round((work+broken)/businesses.length*100)+'%'],
        ['Email Coverage', Math.round(withEmail/businesses.length*100)+'%'],
        ['Phone Coverage', Math.round(withPhone/businesses.length*100)+'%'],
    ];
    let r=3; for(const [l,v] of data) { ws.getCell('A'+r).value=l; ws.getCell('B'+r).value=v; r++; }

    // 2. Verified Directory
    const wd = wb.addWorksheet('Verified Directory');
    wd.columns = [
        {header:'#',key:'id',width:4},{header:'Business Name',key:'name',width:35},
        {header:'Category',key:'cat',width:22},{header:'Email',key:'email',width:35},
        {header:'Phone',key:'phone',width:16},{header:'Website',key:'web',width:42},
        {header:'Status',key:'ws',width:12},{header:'Est.Value',key:'val',width:10},
        {header:'Source',key:'src',width:14},{header:'Address',key:'addr',width:45}
    ];
    wd.getRow(1).eachCell(c=>c.style=h);
    let i=1;
    for(const b of businesses) {
        const w = b.verifiedWebsite||b.website||'';
        wd.addRow({id:i++,name:b.name,cat:b.category||'',email:b.email||'',phone:b.phone||b.mobile||'',
            web:w,ws:w?(b.websiteWorking===true?'Working':'Unknown'):'No Website',
            val:350,src:b.verifiedWebsite&&b.website!==b.verifiedWebsite?'Google Maps':'Original',
            addr:b.address||''});
    }

    // 3. No Website (sales targets)
    const nw = wb.addWorksheet('No Website (Targets)');
    nw.columns = [{header:'#',key:'id',width:4},{header:'Business Name',key:'name',width:35},
        {header:'Category',key:'cat',width:22},{header:'Phone',key:'phone',width:16},
        {header:'Email',key:'email',width:35},{header:'Est.Value',key:'val',width:10}];
    nw.getRow(1).eachCell(c=>c.style=h2);
    let ni=1;
    for(const b of businesses) {
        if(!(b.verifiedWebsite||b.website)) nw.addRow({id:ni++,name:b.name,cat:b.category||'',
            phone:b.phone||b.mobile||'',email:b.email||'',val:'$350-1,500'});
    }

    // 4. Categories
    const cs = wb.addWorksheet('Categories');
    cs.columns = [{header:'Category',key:'cat',width:30},{header:'Total',key:'count',width:8},
        {header:'No Website',key:'nw',width:10},{header:'Working',key:'wk',width:8},
        {header:'Est.Value',key:'val',width:14}];
    cs.getRow(1).eachCell(c=>c.style=h);
    for(const [cat,count] of sc) {
        const items = businesses.filter(x=>(x.category||'')===cat);
        const n = items.filter(x=>!(x.verifiedWebsite||x.website)).length;
        const w = items.filter(x=>x.websiteWorking===true).length;
        cs.addRow({cat,count,nw:n,wk:w,val:'$'+(count*350).toLocaleString()});
    }

    // 5. Pipeline (from localStorage if available)
    const pl = wb.addWorksheet('Pipeline');
    pl.columns = [{header:'Stage',key:'stage',width:18},{header:'Count',key:'count',width:8},{header:'Est.Value',key:'val',width:14}];
    pl.getRow(1).eachCell(c=>c.style=h);
    const stages = ['New','Called','Contacted','Interested','Follow Up','Not Interested','Closed Won','Closed Lost'];
    let sv = {};
    try { sv = JSON.parse(fs.readFileSync('crm-backup.json','utf8')||'{}'); } catch(e) {}
    for(const s of stages) {
        const c = Object.values(sv).filter(x=>x.status===s).length;
        pl.addRow({stage:s,count:c||0,val:''});
    }

    // 6. Service Assessment
    const sa = wb.addWorksheet('Service Assessment');
    sa.columns = [{header:'Service',key:'svc',width:22},{header:'Price',key:'price',width:8},
        {header:'Leads (No Site)',key:'leads',width:14},{header:'Potential Revenue',key:'rev',width:18}];
    sa.getRow(1).eachCell(c=>c.style=h2);
    const prices = {Website:350,'E-commerce':800,'Booking System':600,'POS System':500,CRM:400,
        'Inventory Mgmt':450,'Accounting/ERP':700,'Marketing Automation':350,Communication:300,
        'Online Ordering':500,'HR & Payroll':400,'Analytics & Reporting':300,'Security Systems':500,
        'Cloud Backup':200,'Social Media Mgmt':250,'SMS/Notification':200,'Payment Gateway':300,
        'Delivery Logistics':400,'Learning Mgmt':500,'Patient Mgmt':600};
    for(const [svc,price] of Object.entries(prices)) {
        sa.addRow({svc,price:'$'+price,leads:noWeb,rev:'$'+(noWeb*price).toLocaleString()});
    }

    await wb.xlsx.writeFile('Kigali_Master.xlsx');
    console.log('Master Excel created!');
    console.log('Sheets: Summary, Verified Directory, No Website (Targets), Categories, Pipeline, Service Assessment');
    console.log('Total businesses: '+businesses.length+' | No Website: '+noWeb+' | Working: '+work);
})();
