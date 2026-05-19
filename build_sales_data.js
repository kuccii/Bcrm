const fs = require('fs');
const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));

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

const pullLists = {
    'hot-leads': { name: 'Hot Leads - No Website', desc: 'Businesses with no website + reachable by phone. Highest priority.', cats: null, filter: d => !(d.verifiedWebsite || d.website) && ((d.phone && d.phone.length > 0) || (d.mobile && d.mobile.length > 0)) },
    'restaurants': { name: 'Restaurants & Food', desc: 'Restaurants, cafes, caterers, food retailers without a website.', cats: ['Restaurants', 'Food Retailers', 'Catering', 'Cafes'], filter: d => ['Restaurants', 'Food Retailers', 'Catering', 'Cafes'].includes(d.category || '') && !(d.verifiedWebsite || d.website) },
    'hotels': { name: 'Hotels & Accommodation', desc: 'Hotels, B&Bs, apartments, holiday homes.', cats: ['Hotels', 'Bed and Breakfast', 'Specialist Accommodation', 'Apartments', 'Holiday Homes'], filter: d => ['Hotels', 'Bed and Breakfast', 'Specialist Accommodation', 'Apartments', 'Holiday Homes'].includes(d.category || '') },
    'retail': { name: 'Retail & E-commerce', desc: 'Shops, fashion, furniture, electronics without online presence.', cats: ['Fashion', 'Retail Services', 'Shops', 'Clothing and Accessories', 'Hardware Stores', 'Electronic Equipment', 'Electrical Goods', 'Furniture', 'Beauty Products', 'Gifts', 'Jewellery', 'Books', 'Bicycles'], filter: d => ['Fashion', 'Retail Services', 'Shops', 'Clothing and Accessories', 'Hardware Stores', 'Electronic Equipment', 'Electrical Goods', 'Furniture', 'Beauty Products', 'Gifts', 'Jewellery', 'Books', 'Bicycles'].includes(d.category || '') },
    'medical': { name: 'Medical & Health', desc: 'Doctors, clinics, pharmacies, optical shops.', cats: ['Doctors and Clinics', 'Pharmacies', 'Medical Equipment', 'Opticians', 'Mental Health Care', 'Nursing and Care'], filter: d => ['Doctors and Clinics', 'Pharmacies', 'Medical Equipment', 'Opticians', 'Mental Health Care', 'Nursing and Care'].includes(d.category || '') },
    'tourism': { name: 'Tourism & Travel', desc: 'Tour operators, travel agents, car rental, attractions.', cats: ['Tour Operators', 'Travel Agents', 'Tourist Information', 'Tourism', 'Attractions', 'Car Rental', 'Air Travel', 'Air Transport'], filter: d => ['Tour Operators', 'Travel Agents', 'Tourist Information', 'Tourism', 'Attractions', 'Car Rental', 'Air Travel', 'Air Transport'].includes(d.category || '') },
    'professionals': { name: 'Professional Services', desc: 'Lawyers, consultants, accountants, architects, engineers.', cats: ['Lawyers', 'Consultants', 'Consulting', 'Audit and Accounting', 'Architectural Services', 'Engineering', 'Tax Consultants', 'Legal Services'], filter: d => ['Lawyers', 'Consultants', 'Consulting', 'Audit and Accounting', 'Architectural Services', 'Engineering', 'Tax Consultants', 'Legal Services'].includes(d.category || '') },
    'beauty': { name: 'Beauty & Wellness', desc: 'Salons, spas, fitness centers, beauty shops.', cats: ['Beauty Professionals', 'Beauty Products', 'Hairdressers', 'Fitness', 'Sports'], filter: d => ['Beauty Professionals', 'Beauty Products', 'Hairdressers', 'Fitness', 'Sports'].includes(d.category || '') },
    'automotive': { name: 'Automotive', desc: 'Car dealers, parts, rental, service centers.', cats: ['Automotive', 'Car Parts and Accessories', 'Car Rental', 'Vehicle Services', 'Vehicle Manufacturers'], filter: d => ['Automotive', 'Car Parts and Accessories', 'Car Rental', 'Vehicle Services', 'Vehicle Manufacturers'].includes(d.category || '') },
    'no-website': { name: 'ALL - No Website', desc: 'Every business without a website. Complete list.', cats: null, filter: d => !(d.verifiedWebsite || d.website) },
    'easy-reach': { name: 'Easiest to Reach', desc: 'No website + has email + has phone. Call immediately!', cats: null, filter: d => !(d.verifiedWebsite || d.website) && d.email && d.email.length > 0 && ((d.phone && d.phone.length > 0) || (d.mobile && d.mobile.length > 0)) },
};

// Flatten: each lead gets its list membership + services
const leads = [];
businesses.forEach((b, i) => {
    const id = i + 1;
    const website = b.verifiedWebsite || b.website || '';
    const allLists = [];
    for (const [key, list] of Object.entries(pullLists)) {
        if (list.filter(b)) allLists.push(key);
    }
    // Service recommendation
    const services = getServices(b.category || '');
    const totalValue = services.reduce((sum, s) => sum + ({ Website: 350, 'E-commerce': 800, 'Booking System': 600, 'POS System': 500, CRM: 400, 'Inventory Mgmt': 450, 'Accounting/ERP': 700, 'Marketing Automation': 350, Communication: 300, 'Online Ordering': 500, 'HR & Payroll': 400, 'Analytics & Reporting': 300, 'Security Systems': 500, 'Cloud Backup': 200, 'Social Media Mgmt': 250, 'SMS/Notification': 200, 'Payment Gateway': 300, 'Delivery Logistics': 400, 'Learning Mgmt': 500, 'Patient Mgmt': 600 }[s] || 300), 0);

    leads.push({
        id, name: b.name, cat: b.category || '', email: b.email || '', phone: b.phone || b.mobile || '',
        website, ws: b.websiteWorking === true ? 'working' : (website ? 'broken' : 'none'),
        address: (b.address || '').substring(0, 60), url: b.url || '',
        services: services.join(', '), totalValue, lists: allLists
    });
});

// Compute list stats
const listStats = {};
for (const [key, list] of Object.entries(pullLists)) {
    const items = leads.filter(l => l.lists.includes(key));
    listStats[key] = {
        count: items.length,
        withEmail: items.filter(l => l.email).length,
        withPhone: items.filter(l => l.phone).length,
        totalValue: items.reduce((s, l) => s + l.totalValue, 0),
        avgValue: items.length ? Math.round(items.reduce((s, l) => s + l.totalValue, 0) / items.length) : 0
    };
}

const header = `const PULL_LISTS = ${JSON.stringify(pullLists)};
const LIST_STATS = ${JSON.stringify(listStats)};
const LEADS = ${JSON.stringify(leads)};`;

fs.writeFileSync('sales-data.js', header);
console.log('Sales data created: ' + leads.length + ' leads across ' + Object.keys(pullLists).length + ' lists');
console.log('');
for (const [key, list] of Object.entries(pullLists)) {
    console.log(list.name + ': ' + listStats[key].count + ' leads | $' + listStats[key].totalValue.toLocaleString() + ' est. value');
}