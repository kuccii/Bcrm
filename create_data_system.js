const ExcelJS = require('exceljs');
const fs = require('fs');

(async function() {
    const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Business Scanner';
    wb.created = new Date();
    
    const headerStyle = {
        font: { bold: true, color: { argb: 'FFFFFF' }, size: 11 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } },
        border: { bottom: { style: 'thin', color: { argb: '1B5E20' } } },
        alignment: { horizontal: 'center' }
    };
    
    const titleStyle = {
        font: { bold: true, size: 14, color: { argb: '1B5E20' } },
        alignment: { horizontal: 'left' }
    };
    
    const statsBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } };
    
    // Summary Sheet
    const wsSummary = wb.addWorksheet('Summary');
    wsSummary.columns = [{ header: 'METRIC', key: 'metric', width: 30 }, { header: 'VALUE', key: 'value', width: 15 }];
    wsSummary.getRow(1).eachCell(c => { c.style = headerStyle; });
    
    let withEmail = 0, withPhone = 0, workingCount = 0, notWorkCount = 0, noWebCount = 0;
    let categories = {};
    
    for (const b of businesses) {
        if (b.email && b.email.length > 0) withEmail++;
        if ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0)) withPhone++;
        
        const website = b.verifiedWebsite || b.website || '';
        if (!website) noWebCount++;
        else if (b.websiteWorking === true) workingCount++;
        else notWorkCount++;
        
        const cat = b.category || 'Uncategorized';
        categories[cat] = (categories[cat] || 0) + 1;
    }
    
    const sortedCats = Object.entries(categories).sort((a,b) => b[1] - a[1]);
    
    wsSummary.addRow({ metric: 'KIGALI BUSINESS DIRECTORY', value: '' });
    wsSummary.mergeCells('A1:B1');
    wsSummary.getRow(1).style = titleStyle;
    
    wsSummary.addRow({ metric: '', value: '' });
    wsSummary.addRow({ metric: 'TOTAL BUSINESSES', value: businesses.length });
    wsSummary.addRow({ metric: 'With Email', value: withEmail });
    wsSummary.addRow({ metric: 'With Phone', value: withPhone });
    wsSummary.addRow({ metric: 'Working Websites', value: workingCount });
    wsSummary.addRow({ metric: 'Non-Working Websites', value: notWorkCount });
    wsSummary.addRow({ metric: 'No Website Found', value: noWebCount });
    wsSummary.addRow({ metric: '', value: '' });
    wsSummary.addRow({ metric: 'Email Coverage', value: Math.round(withEmail/businesses.length*100) + '%' });
    wsSummary.addRow({ metric: 'Phone Coverage', value: Math.round(withPhone/businesses.length*100) + '%' });
    wsSummary.addRow({ metric: 'Website Coverage', value: Math.round((workingCount+notWorkCount)/businesses.length*100) + '%' });
    wsSummary.addRow({ metric: 'Total Categories', value: sortedCats.length });
    
    for (let i = 3; i <= 13; i++) {
        wsSummary.getRow(i).eachCell(c => { c.fill = statsBg; });
    }
    
    // Categories Overview
    const wsCats = wb.addWorksheet('All Categories');
    wsCats.columns = [{ header: 'Category', key: 'cat', width: 35 }, { header: 'Count', key: 'count', width: 10 }, { header: '%', key: 'pct', width: 8 }];
    wsCats.getRow(1).style = headerStyle;
    for (const [cat, count] of sortedCats) wsCats.addRow({ cat: cat, count: count, pct: Math.round(count/businesses.length*100) + '%' });
    
    // Main Directory
    const wsMain = wb.addWorksheet('All Businesses');
    wsMain.columns = [
        { header: 'ID', key: 'id', width: 5 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Email', key: 'email', width: 35 },
        { header: 'Phone', key: 'phone', width: 18 },
        { header: 'Website', key: 'website', width: 40 },
        { header: 'Website Status', key: 'websiteStatus', width: 14 },
        { header: 'Address', key: 'address', width: 45 },
        { header: 'Source URL', key: 'url', width: 50 }
    ];
    wsMain.getRow(1).style = headerStyle;
    
    // No Website
    const wsNoWeb = wb.addWorksheet('No Website');
    wsNoWeb.columns = wsMain.columns;
    wsNoWeb.getRow(1).style = headerStyle;
    
    // Non-Working
    const wsNotWork = wb.addWorksheet('Non-Working Websites');
    wsNotWork.columns = wsMain.columns;
    wsNotWork.getRow(1).style = headerStyle;
    
    let mainId = 1, noWebId = 1, notWorkId = 1;
    
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || '';
        const cat = b.category || 'Uncategorized';
        let status = '';
        
        if (!website) { 
            status = 'No Website'; 
            noWebCount--; // Reset and recount below
            wsNoWeb.addRow({ id: noWebId++, name: b.name, category: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: '', websiteStatus: 'No Website', address: b.address || '', url: b.url || '' }); 
        }
        else if (b.websiteWorking === true) { status = 'Working'; }
        else { status = 'Not Working'; notWorkCount--; wsNotWork.addRow({ id: notWorkId++, name: b.name, category: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, websiteStatus: 'Not Working', address: b.address || '', url: b.url || '' }); }
        
        wsMain.addRow({ id: mainId++, name: b.name, category: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, websiteStatus: status, address: b.address || '', url: b.url || '' });
    }
    
    // Create ALL category sheets
    for (const [cat, count] of sortedCats) {
        let safeName = cat.replace(/[\\/:*?"<>|]/g, ' ').substring(0, 31);
        if (safeName.length === 0) safeName = 'Unknown';
        
        const wsCat = wb.addWorksheet(safeName);
        wsCat.columns = wsMain.columns;
        wsCat.getRow(1).style = headerStyle;
        
        let catId = 1;
        for (const b of businesses) {
            if ((b.category || 'Uncategorized') === cat) {
                const website = b.verifiedWebsite || b.website || '';
                let status = website ? (b.websiteWorking === true ? 'Working' : 'Not Working') : 'No Website';
                wsCat.addRow({ id: catId++, name: b.name, category: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, websiteStatus: status, address: b.address || '', url: b.url || '' });
            }
        }
    }
    
    await wb.xlsx.writeFile('Kigali_Data_System.xlsx');
    console.log('Complete! Total Categories:', sortedCats.length);
    console.log('All businesses:', businesses.length);
})();