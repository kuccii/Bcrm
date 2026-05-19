const ExcelJS = require('exceljs');
const fs = require('fs');

(async function() {
    const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    
    const wb = new ExcelJS.Workbook();
    
    const hdr = {
        font: { bold: true, color: { argb: 'FFFFFF' }, size: 10 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1565C0' } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
    };

    const cols = [
        { header: '#', key: 'id', width: 4 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Email', key: 'email', width: 35 },
        { header: 'Phone', key: 'phone', width: 16 },
        { header: 'Current Website', key: 'web', width: 38 },
        { header: 'Sales Pitch', key: 'pitch', width: 60 }
    ];

    // Build industry service needs (same as assessment)
    const industryNeeds = {
        'Restaurants': ['Website', 'Online Ordering', 'Booking System', 'POS System', 'Inventory Mgmt', 'Marketing Automation'],
        'Hotels': ['Website', 'Booking System', 'POS System', 'CRM', 'Marketing Automation', 'Communication'],
        'Tour Operators': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Payment Gateway', 'Communication'],
        'Travel Agents': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Communication', 'Payment Gateway'],
        'Doctors and Clinics': ['Website', 'Patient Mgmt', 'Booking System', 'CRM', 'SMS/Notification', 'Cloud Backup'],
        'Pharmacies': ['Website', 'Inventory Mgmt', 'POS System', 'SMS/Notification', 'Accounting/ERP', 'Online Ordering'],
        'Fashion': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'POS System'],
        'Food Retailers': ['Website', 'Online Ordering', 'POS System', 'Inventory Mgmt', 'Delivery Logistics', 'Marketing Automation'],
        'Retail Services': ['Website', 'E-commerce', 'POS System', 'Inventory Mgmt', 'Accounting/ERP', 'Marketing Automation'],
        'Education': ['Website', 'Learning Mgmt', 'SMS/Notification', 'CRM', 'Cloud Backup', 'Communication'],
        'RealEstate Agents': ['Website', 'CRM', 'Marketing Automation', 'Social Media Mgmt', 'Communication', 'Analytics & Reporting'],
        'Consultants': ['Website', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup', 'Accounting/ERP'],
        'Consulting': ['Website', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup', 'Accounting/ERP'],
        'Lawyers': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP', 'Booking System'],
        'Automotive': ['Website', 'Inventory Mgmt', 'POS System', 'CRM', 'SMS/Notification', 'Accounting/ERP'],
        'Construction': ['Website', 'CRM', 'Accounting/ERP', 'Cloud Backup', 'Analytics & Reporting', 'Communication'],
        'Beauty Professionals': ['Website', 'Booking System', 'Social Media Mgmt', 'Marketing Automation', 'POS System', 'SMS/Notification'],
        'Catering': ['Website', 'Online Ordering', 'Delivery Logistics', 'Inventory Mgmt', 'Marketing Automation', 'Communication'],
        'Cafes': ['Website', 'Online Ordering', 'POS System', 'Inventory Mgmt', 'Marketing Automation', 'Social Media Mgmt'],
        'Hardware Stores': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Accounting/ERP', 'Delivery Logistics'],
    };

    function getServices(cat) {
        return industryNeeds[cat] || ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication', 'Social Media Mgmt'];
    }

    // ======== PULL LIST 1: Hot Leads (no website + has phone) ========
    const ws1 = wb.addWorksheet('Hot Leads - No Website');
    ws1.columns = cols;
    ws1.getRow(1).style = hdr;
    let id1 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        if (!w && ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0))) {
            ws1.addRow({ id: id1++, name: b.name, cat: b.category || '', email: b.email || '', phone: b.phone || b.mobile || '', web: 'NONE', pitch: 'No website at all + reachable by phone. High conversion potential. Offer: Website Package $350' });
        }
    }

    // ======== PULL LIST 2: Restaurants & Food needing website + online ordering ========
    const ws2 = wb.addWorksheet('Restaurants & Food');
    ws2.columns = cols;
    ws2.getRow(1).style = hdr;
    let id2 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Restaurants', 'Food Retailers', 'Catering', 'Cafes'].includes(cat) && !w) {
            ws2.addRow({ id: id2++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: 'NONE', pitch: 'Food business without website. Needs: Website ($350) + Online Ordering ($500) = $850. Pitch: "Your customers cant find your menu online."' });
        }
    }

    // ======== PULL LIST 3: Hotels & Accommodation ========
    const ws3 = wb.addWorksheet('Hotels & Accommodation');
    ws3.columns = cols;
    ws3.getRow(1).style = hdr;
    let id3 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Hotels', 'Bed and Breakfast', 'Specialist Accommodation', 'Apartments', 'Holiday Homes'].includes(cat)) {
            const needs = !w ? 'Needs Website + Booking System ($950)' : 'Upgrade: Booking System + CRM ($1,000)';
            ws3.addRow({ id: id3++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Automate your reservations and never miss a booking."' });
        }
    }

    // ======== PULL LIST 4: Retail & E-commerce ========
    const ws4 = wb.addWorksheet('Retail & E-commerce');
    ws4.columns = cols;
    ws4.getRow(1).style = hdr;
    let id4 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Fashion', 'Retail Services', 'Shops', 'Clothing and Accessories', 'Hardware Stores', 'Electronic Equipment', 'Electrical Goods', 'Furniture', 'Beauty Products', 'Gifts', 'Jewellery', 'Books', 'Bicycles'].includes(cat)) {
            const needs = !w ? 'Needs Website + E-commerce ($1,150)' : 'Add E-commerce + Inventory Mgmt ($1,250)';
            ws4.addRow({ id: id4++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Sell online 24/7 and reach customers who never walk through your door."' });
        }
    }

    // ======== PULL LIST 5: Medical (Doctors, Clinics, Pharmacies) ========
    const ws5 = wb.addWorksheet('Medical & Health');
    ws5.columns = cols;
    ws5.getRow(1).style = hdr;
    let id5 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Doctors and Clinics', 'Pharmacies', 'Medical Equipment', 'Opticians', 'Mental Health Care', 'Nursing and Care'].includes(cat)) {
            const needs = !w ? 'Needs Website + Patient/Booking System ($950)' : 'Add Patient Mgmt + SMS ($800)';
            ws5.addRow({ id: id5++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Let patients book online and get automatic appointment reminders."' });
        }
    }

    // ======== PULL LIST 6: Tourism & Travel ========
    const ws6 = wb.addWorksheet('Tourism & Travel');
    ws6.columns = cols;
    ws6.getRow(1).style = hdr;
    let id6 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Tour Operators', 'Travel Agents', 'Tourist Information', 'Tourism', 'Attractions', 'Car Rental', 'Air Travel', 'Air Transport'].includes(cat)) {
            const needs = !w ? 'Needs Website + Booking System ($950)' : 'Add CRM + Marketing Automation ($750)';
            ws6.addRow({ id: id6++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Automate your tour bookings and client follow-ups."' });
        }
    }

    // ======== PULL LIST 7: Professional Services (Lawyers, Consultants) ========
    const ws7 = wb.addWorksheet('Professional Services');
    ws7.columns = cols;
    ws7.getRow(1).style = hdr;
    let id7 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Lawyers', 'Consultants', 'Consulting', 'Accountants', 'Audit and Accounting', 'Architectural Services', 'Engineering', 'Tax Consultants', 'Legal Services'].includes(cat)) {
            const needs = !w ? 'Needs Website + CRM ($750)' : 'Add CRM + Marketing Automation ($750)';
            ws7.addRow({ id: id7++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Look professional online and never lose a lead again."' });
        }
    }

    // ======== PULL LIST 8: Beauty & Wellness ========
    const ws8 = wb.addWorksheet('Beauty & Wellness');
    ws8.columns = cols;
    ws8.getRow(1).style = hdr;
    let id8 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Beauty Professionals', 'Beauty Products', 'Hairdressers', 'Fitness', 'Sports'].includes(cat)) {
            const needs = !w ? 'Needs Website + Booking System ($950)' : 'Add Booking System + Social Media ($500)';
            ws8.addRow({ id: id8++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Let clients book appointments online 24/7."' });
        }
    }

    // ======== PULL LIST 9: Automotive ========
    const ws9 = wb.addWorksheet('Automotive');
    ws9.columns = cols;
    ws9.getRow(1).style = hdr;
    let id9 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        const cat = b.category || '';
        if (['Automotive', 'Car Parts and Accessories', 'Car Rental', 'Vehicle Services', 'Vehicle Manufacturers'].includes(cat)) {
            const needs = !w ? 'Needs Website + Inventory System ($800)' : 'Add Inventory Mgmt + POS ($950)';
            ws9.addRow({ id: id9++, name: b.name, cat: cat, email: b.email || '', phone: b.phone || b.mobile || '', web: w || 'NONE', pitch: needs + '. Pitch: "Track your inventory and sell parts online."' });
        }
    }

    // ======== PULL LIST 10: All without website (full list) ========
    const ws10 = wb.addWorksheet('ALL - No Website');
    ws10.columns = cols;
    ws10.getRow(1).style = hdr;
    let id10 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        if (!w) {
            ws10.addRow({ id: id10++, name: b.name, cat: b.category || '', email: b.email || '', phone: b.phone || b.mobile || '', web: 'NONE', pitch: 'No website. Offer complete digital package. Est. value: $350-$1,500 depending on industry.' });
        }
    }

    // ======== PULL LIST 11: Have email + phone + no website (easiest to reach) ========
    const ws11 = wb.addWorksheet('Easiest to Reach');
    ws11.columns = cols;
    ws11.getRow(1).style = hdr;
    let id11 = 1;
    for (const b of businesses) {
        const w = b.verifiedWebsite || b.website || '';
        if (!w && b.email && b.email.length > 0 && ((b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0))) {
            ws11.addRow({ id: id11++, name: b.name, cat: b.category || '', email: b.email, phone: b.phone || b.mobile, web: 'NONE', pitch: 'Has both email + phone + no website. WARMEST LEAD. Call immediately.' });
        }
    }

    await wb.xlsx.writeFile('Kigali_Pull_Lists.xlsx');
    console.log('PULL LISTS CREATED: Kigali_Pull_Lists.xlsx');
    console.log('');
    console.log('1. Hot Leads - No Website: ' + (id1-1) + ' leads');
    console.log('2. Restaurants & Food: ' + (id2-1) + ' leads');
    console.log('3. Hotels & Accommodation: ' + (id3-1) + ' leads');
    console.log('4. Retail & E-commerce: ' + (id4-1) + ' leads');
    console.log('5. Medical & Health: ' + (id5-1) + ' leads');
    console.log('6. Tourism & Travel: ' + (id6-1) + ' leads');
    console.log('7. Professional Services: ' + (id7-1) + ' leads');
    console.log('8. Beauty & Wellness: ' + (id8-1) + ' leads');
    console.log('9. Automotive: ' + (id9-1) + ' leads');
    console.log('10. ALL - No Website: ' + (id10-1) + ' leads');
    console.log('11. Easiest to Reach: ' + (id11-1) + ' leads (email + phone + no website)');
})();