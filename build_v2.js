const ExcelJS = require('exceljs');
const fs = require('fs');

(async function() {
    const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Business Scanner';
    wb.created = new Date();
    
    const greenHeader = { font: { bold: true, color: { argb: 'FFFFFF' }, size: 11 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } }, alignment: { horizontal: 'center' } };
    const titleStyle = { font: { bold: true, size: 16, color: { argb: '1B5E20' } } };
    
    let withEmail = 0, withPhone = 0, workingCount = 0, notWorkCount = 0, noWebCount = 0;
    let categories = {};
    for (const b of businesses) {
        if (b.email && b.email.length > 0) withEmail++;
        if ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0)) withPhone++;
        const website = b.verifiedWebsite || b.website || '';
        if (!website) noWebCount++;
        else if (b.websiteWorking === true) workingCount++;
        else notWorkCount++;
        categories[b.category || 'Uncategorized'] = (categories[b.category || 'Uncategorized'] || 0) + 1;
    }
    const sortedCats = Object.entries(categories).sort((a,b) => b[1] - a[1]);
    
    const wsSum = wb.addWorksheet('Summary');
    wsSum.mergeCells('A1:B1');
    wsSum.getCell('A1').value = 'KIGALI BUSINESS DIRECTORY (Google Verified)';
    wsSum.getCell('A1').style = titleStyle;
    wsSum.addRow({});
    const stats = [
        ['TOTAL BUSINESSES', businesses.length],
        ['With Website (Google Verified)', workingCount + notWorkCount],
        ['Working Websites', workingCount],
        ['No Website Found', noWebCount],
        ['With Email', withEmail],
        ['With Phone', withPhone],
        ['Total Categories', sortedCats.length],
        ['', ''],
        ['Email Coverage', Math.round(withEmail/businesses.length*100) + '%'],
        ['Phone Coverage', Math.round(withPhone/businesses.length*100) + '%'],
        ['Website Coverage', Math.round((workingCount+notWorkCount)/businesses.length*100) + '%'],
    ];
    let r = 3;
    for (const [l, v] of stats) { wsSum.getCell('A' + r).value = l; wsSum.getCell('B' + r).value = v; r++; }
    
    const wsMain = wb.addWorksheet('All Businesses');
    wsMain.columns = [
        { header: 'ID', key: 'id', width: 5 }, { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 }, { header: 'Email', key: 'email', width: 35 },
        { header: 'Phone', key: 'phone', width: 16 }, { header: 'Website', key: 'website', width: 42 },
        { header: 'Status', key: 'ws', width: 14 }, { header: 'Source', key: 'src', width: 16 },
        { header: 'Address', key: 'addr', width: 48 }
    ];
    wsMain.getRow(1).style = greenHeader;
    
    const wsNoWeb = wb.addWorksheet('No Website');
    wsNoWeb.columns = wsMain.columns;
    wsNoWeb.getRow(1).style = greenHeader;
    
    const wsVerified = wb.addWorksheet('Google Verified');
    wsVerified.columns = wsMain.columns;
    wsVerified.getRow(1).style = greenHeader;
    
    let mainId = 1, noWebId = 1, verifiedId = 1;
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        const website = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        const isGoogle = b.verifiedWebsite && b.website !== b.verifiedWebsite;
        const status = website ? (b.websiteWorking === true ? 'Working' : 'Unknown') : 'No Website';
        const source = isGoogle ? 'Google Places' : b.googleVerified ? 'Google' : 'Original';
        
        wsMain.addRow({ id: mainId++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, ws: status, src: source, addr: b.address || '' });
        if (i % 2 === 0) wsMain.getRow(i + 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F5F5' } };
        
        if (!website) {
            wsNoWeb.addRow({ id: noWebId++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: '', ws: 'No Website', src: '', addr: b.address || '' });
            wsNoWeb.getRow(noWebId).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBEE' } };
        }
        
        if (isGoogle) {
            wsVerified.addRow({ id: verifiedId++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, ws: status, src: source, addr: b.address || '' });
        }
    }
    
    const colDefs = [
        { header: '#', key: 'id', width: 4 }, { header: 'Business Name', key: 'name', width: 38 },
        { header: 'Website', key: 'website', width: 42 }, { header: 'Status', key: 'ws', width: 14 },
        { header: 'Contact', key: 'contact', width: 35 }
    ];
    
    for (const [cat, count] of sortedCats.slice(0, 30)) {
        let safeName = cat.replace(/[\\/:*?"<>|]/g, ' ').substring(0, 31).trim();
        if (!safeName) safeName = 'Other';
        const wsCat = wb.addWorksheet(safeName);
        wsCat.columns = colDefs;
        wsCat.getRow(1).style = greenHeader;
        let catId = 1;
        for (const b of businesses) {
            if ((b.category || '') === cat) {
                const website = b.verifiedWebsite || b.website || '';
                wsCat.addRow({ id: catId++, name: b.name, website: website, ws: website ? 'Yes' : 'No', contact: b.email || b.phone || b.mobile || '' });
            }
        }
    }
    
    await wb.xlsx.writeFile('Kigali_Data_System_v2.xlsx');
    console.log('Saved: Kigali_Data_System_v2.xlsx');
    console.log('Businesses:', businesses.length, '| With Website:', workingCount + notWorkCount, '| No Website:', noWebCount);
})();
