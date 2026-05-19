const ExcelJS = require('exceljs');
const fs = require('fs');

(async function() {
    const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Sales Prospecting System';
    wb.created = new Date();
    
    const hdrStyle = {
        font: { bold: true, color: { argb: 'FFFFFF' }, size: 10 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1565C0' } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: { bottom: { style: 'medium', color: { argb: '0D47A1' } } }
    };
    
    const greenHdr = { ...hdrStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } } };
    const orangeHdr = { ...hdrStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E65100' } } };
    const redHdr = { ...hdrStyle, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C62828' } } };

    // ======== PREP DATA ========
    let withEmail = 0, withPhone = 0, workingCount = 0, notWorkCount = 0, noWebCount = 0;
    let categories = {};
    
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || '';
        if (!website) noWebCount++;
        else if (b.websiteWorking === true) workingCount++;
        else notWorkCount++;
        if (b.email && b.email.length > 0) withEmail++;
        if ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0)) withPhone++;
        categories[b.category || 'Uncategorized'] = (categories[b.category || 'Uncategorized'] || 0) + 1;
    }
    
    const sortedCats = Object.entries(categories).sort((a,b) => b[1] - a[1]);

    // ======== 1. SALES DASHBOARD ========
    const wsDash = wb.addWorksheet('Sales Dashboard');
    wsDash.columns = [{ header: '', key: 'l', width: 32 }, { header: '', key: 'v', width: 15 }];
    
    wsDash.mergeCells('A1:B1');
    wsDash.getCell('A1').value = 'SALES PROSPECTING DASHBOARD';
    wsDash.getCell('A1').style = { font: { bold: true, size: 18, color: { argb: '0D47A1' } }, alignment: { horizontal: 'center' } };
    wsDash.getRow(1).height = 30;
    
    wsDash.mergeCells('A3:B3');
    wsDash.getCell('A3').value = 'MARKET OVERVIEW';
    wsDash.getCell('A3').style = { font: { bold: true, size: 13, color: { argb: '1565C0' } } };
    
    const rows = [
        ['Total Businesses', businesses.length],
        ['Need Website (HIGH PRIORITY)', noWebCount],
        ['Website Not Working (MEDIUM PRIORITY)', notWorkCount],
        ['Have Working Website', workingCount],
        ['', ''],
        ['Reachable by Email', withEmail],
        ['Reachable by Phone', withPhone],
        ['Total Serviceable Market', businesses.length],
        ['', ''],
        ['PRICING STRATEGY (per business):', ''],
        ['Website Creation Package', '$200 - $500'],
        ['Website Repair Package', '$100 - $250'],
        ['SEO Optimization', '$150 - $300'],
        ['Email & Digital Setup', '$50 - $150'],
        ['Full Digital Package', '$400 - $800'],
        ['', ''],
        ['PROJECTED REVENUE:', ''],
        ['If 10% of "No Website" buy ($350 avg)', '$' + Math.round(noWebCount * 0.1 * 350).toLocaleString()],
        ['If 20% of "No Website" buy ($350 avg)', '$' + Math.round(noWebCount * 0.2 * 350).toLocaleString()],
        ['If 30% of "No Website" buy ($350 avg)', '$' + Math.round(noWebCount * 0.3 * 350).toLocaleString()],
    ];
    
    let r = 4;
    for (const [label, val] of rows) {
        wsDash.getCell('A' + r).value = label;
        wsDash.getCell('B' + r).value = val;
        wsDash.getCell('A' + r).style = label.includes('PRIORITY') ? { font: { bold: true, color: { argb: 'C62828' } } } : {};
        wsDash.getCell('B' + r).style = { alignment: { horizontal: 'center' } };
        r++;
    }

    // ======== 2. LEAD SCORING (PRIORITY LIST) ========
    const wsLeads = wb.addWorksheet('Lead Scoring');
    wsLeads.columns = [
        { header: 'Priority', key: 'pri', width: 6 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Lead Score', key: 'score', width: 8 },
        { header: 'Service to Sell', key: 'service', width: 28 },
        { header: 'Est. Value', key: 'value', width: 12 },
        { header: 'Email', key: 'email', width: 35 },
        { header: 'Phone', key: 'phone', width: 16 },
        { header: 'Current Website', key: 'website', width: 38 },
        { header: 'Sales Stage', key: 'stage', width: 16 },
        { header: 'Notes', key: 'notes', width: 30 }
    ];
    wsLeads.getRow(1).style = hdrStyle;
    wsLeads.getRow(1).height = 25;
    
    // Score each business
    let leads = [];
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || '';
        let score = 0;
        let service = '';
        let estValue = '';
        
        if (!website) { score += 50; service = 'Website Creation'; estValue = '$350'; }
        else if (b.websiteWorking !== true) { score += 30; service = 'Website Repair'; estValue = '$175'; }
        else { score += 10; service = 'SEO / Optimization'; estValue = '$225'; }
        
        if (b.email && b.email.length > 0) score += 20;
        if ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0)) score += 15;
        if (b.category) score += 5;
        
        // Boost for certain categories
        const highValueCats = ['Hotels', 'Restaurants', 'Tour Operators', 'Travel Agents', 'RealEstate Agents', 'Doctors and Clinics', 'Pharmacies'];
        if (highValueCats.includes(b.category)) score += 10;
        
        leads.push({
            name: b.name,
            cat: b.category || '',
            score: score,
            service: service,
            value: estValue,
            email: b.email || '',
            phone: b.phone || b.mobile || '',
            website: website,
            stage: 'New',
            notes: score >= 70 ? 'HOT LEAD - Call today!' : score >= 50 ? 'Warm lead' : 'Cold lead'
        });
    }
    
    // Sort by score descending
    leads.sort((a, b) => b.score - a.score);
    
    for (let i = 0; i < leads.length; i++) {
        const l = leads[i];
        wsLeads.addRow({ pri: i + 1, name: l.name, cat: l.cat, score: l.score, service: l.service, value: l.value, email: l.email, phone: l.phone, website: l.website, stage: l.stage, notes: l.notes });
        
        const rowNum = i + 2;
        if (l.score >= 70) wsLeads.getRow(rowNum).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBEE' } };
        else if (l.score >= 50) wsLeads.getRow(rowNum).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E0' } };
        else if (i % 2 === 0) wsLeads.getRow(rowNum).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F5F5' } };
    }

    // ======== 3. NEED WEBSITE (TOP PRIORITY) ========
    const wsNoWeb = wb.addWorksheet('Priority 1 - Need Website');
    wsNoWeb.columns = [
        { header: '#', key: 'id', width: 4 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Email', key: 'email', width: 38 },
        { header: 'Phone', key: 'phone', width: 16 },
        { header: 'Sales Stage', key: 'stage', width: 16 },
        { header: 'Contacted Date', key: 'date', width: 14 },
        { header: 'Notes', key: 'notes', width: 45 }
    ];
    wsNoWeb.getRow(1).style = redHdr;
    wsNoWeb.getRow(1).height = 25;
    
    let noWebId = 1;
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || '';
        if (!website) {
            wsNoWeb.addRow({ id: noWebId++, name: b.name, cat: b.category || '', email: b.email || '', phone: b.phone || b.mobile || '', stage: 'New', date: '', notes: '' });
            wsNoWeb.getRow(noWebId).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBEE' } };
        }
    }
    
    // ======== 4. WEBSITE NOT WORKING (MEDIUM PRIORITY) ========
    const wsNotWork = wb.addWorksheet('Priority 2 - Website Issues');
    wsNotWork.columns = wsNoWeb.columns;
    wsNotWork.getRow(1).style = orangeHdr;
    wsNotWork.getRow(1).height = 25;
    
    let notWorkId = 1;
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || '';
        if (website && b.websiteWorking !== true) {
            wsNotWork.addRow({ id: notWorkId++, name: b.name, cat: b.category || '', email: b.email || '', phone: b.phone || b.mobile || '', stage: 'New', date: '', notes: '' });
            wsNotWork.getRow(notWorkId).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E0' } };
        }
    }
    
    // ======== 5. HIGH VALUE CATEGORIES ========
    const wsHV = wb.addWorksheet('Priority 3 - High Value Cats');
    wsHV.columns = [
        { header: '#', key: 'id', width: 4 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Email', key: 'email', width: 38 },
        { header: 'Phone', key: 'phone', width: 16 },
        { header: 'Website', key: 'website', width: 38 },
        { header: 'Website Status', key: 'status', width: 14 },
        { header: 'Notes', key: 'notes', width: 45 }
    ];
    wsHV.getRow(1).style = hdrStyle;
    wsHV.getRow(1).height = 25;
    
    const highValueCats = ['Hotels', 'Restaurants', 'Tour Operators', 'Travel Agents', 'Doctors and Clinics', 'Pharmacies', 'RealEstate Agents', 'Consultants', 'Lawyers'];
    let hvId = 1;
    for (const b of businesses) {
        if (highValueCats.includes(b.category || '')) {
            const website = b.verifiedWebsite || b.website || '';
            let status = website ? (b.websiteWorking === true ? 'Working' : 'Not Working') : 'No Website';
            wsHV.addRow({ id: hvId++, name: b.name, cat: b.category, email: b.email || '', phone: b.phone || b.mobile || '', website: website, status: status, notes: '' });
            if (hvId % 2 === 0) wsHV.getRow(hvId).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E3F2FD' } };
        }
    }
    
    // ======== 6. OUTREACH SCRIPTS ========
    const wsScripts = wb.addWorksheet('Sales Scripts');
    wsScripts.columns = [{ header: 'Scenario', key: 's', width: 25 }, { header: 'Script', key: 'script', width: 100 }];
    wsScripts.getRow(1).style = hdrStyle;
    wsScripts.getRow(1).height = 25;
    
    const scripts = [
        ['No Website - Phone Script',
`"Hello [Name], this is [Your Name] from [Company]. I noticed [Business Name] doesn't have a website yet.
In today's market, 80% of customers search online before buying. We create professional websites starting at $200.
Can I show you some examples specific to the [Category] industry?"`],
        ['No Website - Email Template',
`Subject: Help [Business Name] get online

Hi [Name],

I noticed [Business Name] doesn't have an active website. In Rwanda, more customers are searching online every day.

We build professional websites for businesses like yours starting from $200.
We handle everything: domain, hosting, design, and mobile optimization.

Would you be open to a quick 10-minute call to discuss?

Best regards,
[Your Name]
[Phone Number]`],
        ['Website Not Working - Phone Script',
`"Hi [Name], this is [Your Name]. I was looking at [Business Name]'s website at [URL] and noticed it seems to be down/having issues.

A broken website costs you customers every day. We can fix it fast, usually within 24-48 hours, starting at $100.

Can I take a quick look and give you a free assessment?"`],
        ['Website Not Working - Email',
`Subject: Your website at [URL] needs attention

Hi [Name],

I visited your website at [URL] and noticed it wasn't loading properly. This could be costing you potential customers.

We specialize in fixing website issues quickly. I can have a specialist review and provide a fix estimate within 24 hours.

Let me know if you'd like a free website audit.

Best,
[Your Name]`],
        ['Have Website - SEO/Upgrade Pitch',
`"Hi [Name], I was checking out [Business Name]'s website at [URL]. It looks good, but I noticed a few opportunities to improve it.

We can optimize it to show up higher in Google searches, add new features, or modernize the design. Our SEO packages start at $150.

Interested in a free website audit?"`],
        ['Follow-up Script (No response)',
`"Hi [Name], I reached out last week about your website. Many businesses in [Category] have seen 40%+ more customers after getting online.

I'd love to help [Business Name] grow. Can we schedule a quick 5-minute call?"`]
    ];
    
    for (const [scenario, script] of scripts) {
        wsScripts.addRow({ s: scenario, script: script });
    }
    
    // ======== 7. TRACKING SHEET ========
    const wsTrack = wb.addWorksheet('My Sales Tracker');
    wsTrack.columns = [
        { header: 'Date', key: 'date', width: 12 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Phone', key: 'phone', width: 16 },
        { header: 'Email', key: 'email', width: 35 },
        { header: 'Service Offered', key: 'service', width: 22 },
        { header: 'Status', key: 'status', width: 14 },
        { header: 'Follow-up Date', key: 'follow', width: 14 },
        { header: 'Notes', key: 'notes', width: 40 }
    ];
    wsTrack.getRow(1).style = hdrStyle;
    wsTrack.getRow(1).height = 25;
    
    // ======== 8. CATEGORY CAMPAIGNS ========
    const wsCamp = wb.addWorksheet('Campaigns by Category');
    wsCamp.columns = [
        { header: 'Category', key: 'cat', width: 30 },
        { header: 'Businesses', key: 'count', width: 10 },
        { header: 'Need Website', key: 'need', width: 10 },
        { header: 'Have Email', key: 'email', width: 10 },
        { header: 'Have Phone', key: 'phone', width: 10 },
        { header: 'Strategy', key: 'strategy', width: 65 }
    ];
    wsCamp.getRow(1).style = hdrStyle;
    wsCamp.getRow(1).height = 25;
    
    for (const [cat, count] of sortedCats) {
        let needWeb = 0, haveEmail = 0, havePhone = 0;
        for (const b of businesses) {
            if ((b.category || 'Uncategorized') === cat) {
                const website = b.verifiedWebsite || b.website || '';
                if (!website) needWeb++;
                if (b.email && b.email.length > 0) haveEmail++;
                if ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0)) havePhone++;
            }
        }
        
        let strategy = '';
        if (needWeb > count * 0.5) strategy = 'Bulk outreach: "Most ' + cat + ' in Kigali now have websites. Dont get left behind."';
        else if (needWeb > 0) strategy = 'Targeted: "We noticed your competitors in ' + cat + ' are online. Lets get you there too."';
        else strategy = 'Upsell: "Your website works great! Want to rank higher on Google?"';
        
        wsCamp.addRow({ cat: cat, count: count, need: needWeb, email: haveEmail, phone: havePhone, strategy: strategy });
    }
    
    // ======== 9. CATEGORY SHEETS ========
    const colDefs = [
        { header: '#', key: 'id', width: 4 },
        { header: 'Business Name', key: 'name', width: 38 },
        { header: 'Email', key: 'email', width: 38 },
        { header: 'Phone', key: 'phone', width: 18 },
        { header: 'Website Status', key: 'status', width: 14 },
        { header: 'Sales Stage', key: 'stage', width: 14 },
        { header: 'Est. Value', key: 'value', width: 12 },
        { header: 'Notes', key: 'notes', width: 35 }
    ];
    
    for (const [cat, count] of sortedCats.slice(0, 20)) {
        let safeName = cat.replace(/[\\/:*?"<>|]/g, ' ').substring(0, 31).trim();
        if (!safeName) safeName = 'Other';
        
        const wsCat = wb.addWorksheet(safeName);
        wsCat.columns = colDefs;
        wsCat.getRow(1).style = hdrStyle;
        wsCat.getRow(1).height = 25;
        
        let catId = 1;
        for (const b of businesses) {
            if ((b.category || 'Uncategorized') === cat) {
                const website = b.verifiedWebsite || b.website || '';
                let service = '';
                if (!website) service = 'Website ($350)';
                else if (b.websiteWorking !== true) service = 'Repair ($175)';
                else service = 'SEO ($225)';
                
                wsCat.addRow({ id: catId++, name: b.name, email: b.email || '', phone: b.phone || b.mobile || '', status: website ? (b.websiteWorking === true ? 'Working' : 'Issues') : 'No Site', stage: 'New', value: service, notes: '' });
                if (catId % 2 === 0) wsCat.getRow(catId).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F5F5' } };
            }
        }
    }
    
    await wb.xlsx.writeFile('Kigali_Sales_Prospecting.xlsx');
    console.log('SALES SYSTEM COMPLETE!');
    console.log('File: Kigali_Sales_Prospecting.xlsx');
    console.log('');
    console.log('Sheets:');
    console.log('  1. Sales Dashboard - Overview & revenue projections');
    console.log('  2. Lead Scoring - All 2,750 businesses ranked by priority');
    console.log('  3. Priority 1 - Need Website (' + noWebCount + ' leads - HIGH VALUE)');
    console.log('  4. Priority 2 - Website Issues (' + notWorkCount + ' leads)');
    console.log('  5. Priority 3 - High Value Categories (' + hvId + ' leads)');
    console.log('  6. Sales Scripts - Ready-to-use call & email templates');
    console.log('  7. My Sales Tracker - Track your outreach progress');
    console.log('  8. Campaigns by Category - Strategy for each category');
    console.log('  9. Top 20 Category sheets - Targeted sales lists');
    console.log('');
    console.log('Estimated Revenue Potential:');
    console.log('  10% conversion: $' + Math.round(noWebCount * 0.1 * 350).toLocaleString());
    console.log('  20% conversion: $' + Math.round(noWebCount * 0.2 * 350).toLocaleString());
    console.log('  30% conversion: $' + Math.round(noWebCount * 0.3 * 350).toLocaleString());
})();