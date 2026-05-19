const ExcelJS = require('exceljs');
const fs = require('fs');

(async function() {
    const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Business Scanner';
    wb.created = new Date();
    
    // Beautiful Styles
    const greenHeader = {
        font: { bold: true, color: { argb: 'FFFFFF' }, size: 11 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } },
        border: { bottom: { style: 'medium', color: { argb: '1B5E20' } } },
        alignment: { horizontal: 'center', vertical: 'middle' }
    };
    
    const titleStyle = {
        font: { bold: true, size: 16, color: { argb: '1B5E20' } }
    };
    
    const statLabel = { font: { bold: true, size: 11, color: { argb: '388E3C' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } } };
    const statValue = { font: { size: 11 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F1F8E9' } }, alignment: { horizontal: 'center' } };
    
    // Count data
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
    
    // ============ SUMMARY SHEET ============
    const wsSummary = wb.addWorksheet('Summary');
    wsSummary.columns = [{ header: '', key: 'label', width: 28 }, { header: '', key: 'value', width: 12 }];
    
    // Title
    wsSummary.mergeCells('A1:B1');
    wsSummary.getCell('A1').value = '📊 KIGALI BUSINESS DIRECTORY';
    wsSummary.getCell('A1').style = { font: { bold: true, size: 18, color: { argb: '1B5E20' } }, alignment: { horizontal: 'center' } };
    wsSummary.getRow(1).height = 30;
    
    // Stats
    const stats = [
        { label: '📍 Total Businesses', value: businesses.length },
        { label: '📧 With Email', value: withEmail },
        { label: '📞 With Phone', value: withPhone },
        { label: '✅ Working Websites', value: workingCount },
        { label: '❌ Non-Working Websites', value: notWorkCount },
        { label: '🚫 No Website Found', value: noWebCount },
        { label: '📈 Email Coverage', value: Math.round(withEmail/businesses.length*100) + '%' },
        { label: '📱 Phone Coverage', value: Math.round(withPhone/businesses.length*100) + '%' },
        { label: '🌐 Website Coverage', value: Math.round((workingCount+notWorkCount)/businesses.length*100) + '%' },
        { label: '📂 Total Categories', value: sortedCats.length }
    ];
    
    let row = 3;
    for (const s of stats) {
        wsSummary.getCell('A' + row).value = s.label;
        wsSummary.getCell('B' + row).value = s.value;
        wsSummary.getCell('A' + row).style = statLabel;
        wsSummary.getCell('B' + row).style = statValue;
        row++;
    }
    
    // ============ CATEGORIES OVERVIEW ============
    const wsCats = wb.addWorksheet('All Categories');
    wsCats.columns = [{ header: 'Category', key: 'cat', width: 38 }, { header: 'Count', key: 'count', width: 10 }, { header: '%', key: 'pct', width: 8 }];
    wsCats.getRow(1).style = greenHeader;
    wsCats.getRow(1).height = 25;
    
    for (let i = 0; i < sortedCats.length; i++) {
        const [cat, count] = sortedCats[i];
        wsCats.addRow({ cat: cat, count: count, pct: Math.round(count/businesses.length*100) + '%' });
        if (i % 2 === 0) wsCats.getRow(i + 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F5F5' } };
    }
    
    // ============ MAIN DIRECTORY ============
    const colDefs = [
        { header: '#', key: 'id', width: 5 },
        { header: 'Business Name', key: 'name', width: 38 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Email', key: 'email', width: 38 },
        { header: 'Phone', key: 'phone', width: 16 },
        { header: 'Website', key: 'website', width: 42 },
        { header: 'Status', key: 'websiteStatus', width: 14 },
        { header: 'Address', key: 'address', width: 48 }
    ];
    
    const wsMain = wb.addWorksheet('All Businesses');
    wsMain.columns = colDefs;
    wsMain.getRow(1).style = greenHeader;
    wsMain.getRow(1).height = 25;
    
    // No Website Sheet
    const wsNoWeb = wb.addWorksheet('No Website');
    wsNoWeb.columns = colDefs;
    wsNoWeb.getRow(1).style = greenHeader;
    wsNoWeb.getRow(1).height = 25;
    
    // Non-Working Sheet
    const wsNotWork = wb.addWorksheet('Non-Working Websites');
    wsNotWork.columns = colDefs;
    wsNotWork.getRow(1).style = greenHeader;
    wsNotWork.getRow(1).height = 25;
    
    let mainId = 1, noWebId = 1, notWorkId = 1;
    
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        const website = b.verifiedWebsite || b.website || '';
        const cat = b.category || 'Uncategorized';
        
        let status = '';
        let rowData = { id: mainId++, name: b.name, category: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, websiteStatus: '', address: b.address || '' };
        
        if (!website) {
            status = 'No Website';
            wsNoWeb.addRow({ id: noWebId++, ...rowData });
        } else if (b.websiteWorking === true) {
            status = 'Working';
        } else {
            status = 'Not Working';
            wsNotWork.addRow({ id: notWorkId++, ...rowData, websiteStatus: status });
        }
        
        rowData.websiteStatus = status;
        wsMain.addRow(rowData);
        
        // Alternating row colors
        const fill = i % 2 === 0 ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } } : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F5F5' } };
        wsMain.getRow(i + 2).fill = fill;
    }
    
    // Style No Website and Not Working sheets
    for (let i = 2; i < wsNoWeb.rowCount + 1; i++) {
        wsNoWeb.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBEE' } };
    }
    for (let i = 2; i < wsNotWork.rowCount + 1; i++) {
        wsNotWork.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E0' } };
    }
    
    // ============ ALL CATEGORY SHEETS ============
    for (const [cat, count] of sortedCats) {
        let safeName = cat.replace(/[\\/:*?"<>|]/g, ' ').substring(0, 31);
        if (!safeName.trim()) safeName = 'Unknown';
        
        const wsCat = wb.addWorksheet(safeName);
        wsCat.columns = colDefs;
        wsCat.getRow(1).style = greenHeader;
        wsCat.getRow(1).height = 25;
        
        let catId = 1;
        let catRow = 2;
        
        for (const b of businesses) {
            if ((b.category || 'Uncategorized') === cat) {
                const website = b.verifiedWebsite || b.website || '';
                let status = website ? (b.websiteWorking === true ? 'Working' : 'Not Working') : 'No Website';
                
                wsCat.addRow({ id: catId++, name: b.name, category: cat, email: b.email || '', phone: b.phone || b.mobile || '', website: website, websiteStatus: status, address: b.address || '' });
                
                if (catRow % 2 === 0) wsCat.getRow(catRow).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } };
                catRow++;
            }
        }
    }
    
    await wb.xlsx.writeFile('Kigali_Data_System.xlsx');
    console.log('✅ Complete!');
    console.log('   Total: ' + businesses.length);
    console.log('   Categories: ' + sortedCats.length);
    console.log('   Sheets: Summary, All Categories, All Businesses, No Website, Non-Working Websites, +' + sortedCats.length + ' category sheets');
})();