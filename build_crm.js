const fs = require('fs');
const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));

const industryNeeds = {
  'Restaurants': ['Website','Online Ordering','Booking System','POS System','Inventory Mgmt','Marketing Automation'],
  'Hotels': ['Website','Booking System','POS System','CRM','Marketing Automation','Communication'],
  'Tour Operators': ['Website','Booking System','CRM','Marketing Automation','Payment Gateway','Communication'],
  'Travel Agents': ['Website','Booking System','CRM','Marketing Automation','Communication','Payment Gateway'],
  'Doctors and Clinics': ['Website','Patient Mgmt','Booking System','CRM','SMS/Notification','Cloud Backup'],
  'Pharmacies': ['Website','Inventory Mgmt','POS System','SMS/Notification','Accounting/ERP','Online Ordering'],
  'Fashion': ['Website','E-commerce','Social Media Mgmt','Inventory Mgmt','Marketing Automation','POS System'],
  'Food Retailers': ['Website','Online Ordering','POS System','Inventory Mgmt','Delivery Logistics','Marketing Automation'],
  'Retail Services': ['Website','E-commerce','POS System','Inventory Mgmt','Accounting/ERP','Marketing Automation'],
  'Education': ['Website','Learning Mgmt','SMS/Notification','CRM','Cloud Backup','Communication'],
  'RealEstate Agents': ['Website','CRM','Marketing Automation','Social Media Mgmt','Communication','Analytics & Reporting'],
  'Consultants': ['Website','CRM','Marketing Automation','Communication','Cloud Backup','Accounting/ERP'],
  'Consulting': ['Website','CRM','Marketing Automation','Communication','Cloud Backup','Accounting/ERP'],
  'Lawyers': ['Website','CRM','Cloud Backup','Communication','Accounting/ERP','Booking System'],
  'Automotive': ['Website','Inventory Mgmt','POS System','CRM','SMS/Notification','Accounting/ERP'],
  'Construction': ['Website','CRM','Accounting/ERP','Cloud Backup','Analytics & Reporting','Communication'],
  'Beauty Professionals': ['Website','Booking System','Social Media Mgmt','Marketing Automation','POS System','SMS/Notification'],
  'Catering': ['Website','Online Ordering','Delivery Logistics','Inventory Mgmt','Marketing Automation','Communication'],
  'Cafes': ['Website','Online Ordering','POS System','Inventory Mgmt','Marketing Automation','Social Media Mgmt'],
  'Hardware Stores': ['Website','E-commerce','Inventory Mgmt','POS System','Accounting/ERP','Delivery Logistics'],
};

function getServices(cat) { return industryNeeds[cat] || ['Website','CRM','Marketing Automation','Cloud Backup','Communication','Social Media Mgmt']; }

const svcPrices = {
  'Website':350,'E-commerce':800,'Booking System':600,'POS System':500,'CRM':400,'Inventory Mgmt':450,
  'Accounting/ERP':700,'Marketing Automation':350,'Communication':300,'Online Ordering':500,'HR & Payroll':400,
  'Analytics & Reporting':300,'Security Systems':500,'Cloud Backup':200,'Social Media Mgmt':250,'SMS/Notification':200,
  'Payment Gateway':300,'Delivery Logistics':400,'Learning Mgmt':500,'Patient Mgmt':600
};

const crmData = businesses.map((b, i) => {
  const id = i + 1;
  const website = b.verifiedWebsite || b.website || '';
  const services = getServices(b.category || '');
  const totalValue = services.reduce((s, n) => s + (svcPrices[n] || 300), 0);
  const ws = b.websiteWorking === true ? 'working' : (website ? 'broken' : 'none');
  return {
    id, name: b.name, cat: b.category || '', email: b.email || '',
    phone: b.phone || b.mobile || '', website, ws,
    address: (b.address || '').substring(0, 80),
    services: services.join(', '), totalValue,
    needsWebsite: ws === 'none',
    priority: ws === 'none' ? 'high' : ws === 'broken' ? 'medium' : 'low',
    source: b.verifiedWebsite && b.website !== b.verifiedWebsite ? 'Google Places' : b.googleVerified ? 'Google' : 'Original'
  };
});

const cats = {};
crmData.forEach(d => { cats[d.cat] = (cats[d.cat] || 0) + 1; });
const sortedCats = Object.entries(cats).sort((a,b) => b[1] - a[1]);

// Build CRM_LIST_DATA statically
function count(key, fn) {
  const items = crmData.filter(fn);
  return { count: items.length, totalValue: items.reduce((s, d) => s + d.totalValue, 0) };
}

const foodCats = ['Restaurants','Food Retailers','Catering','Cafes'];
const hotelCats = ['Hotels','Bed and Breakfast','Specialist Accommodation','Apartments','Holiday Homes'];
const retailCats = ['Fashion','Retail Services','Shops','Clothing and Accessories','Hardware Stores','Electronic Equipment','Electrical Goods','Furniture','Beauty Products','Gifts','Jewellery','Books','Bicycles'];
const medicalCats = ['Doctors and Clinics','Pharmacies','Medical Equipment','Opticians','Mental Health Care','Nursing and Care'];
const tourismCats = ['Tour Operators','Travel Agents','Tourist Information','Tourism','Attractions','Car Rental','Air Travel','Air Transport'];
const profCats = ['Lawyers','Consultants','Consulting','Audit and Accounting','Architectural Services','Engineering','Tax Consultants','Legal Services'];
const beautyCats = ['Beauty Professionals','Beauty Products','Hairdressers','Fitness','Sports'];
const autoCats = ['Automotive','Car Parts and Accessories','Car Rental','Vehicle Services','Vehicle Manufacturers'];

const listMeta = {
  'hot-leads': { name: 'Hot Leads \u2014 No Website', icon: '\uD83D\uDD25', filter: 'hot' },
  'restaurants': { name: 'Restaurants & Food', icon: '\uD83C\uDF7D', filter: 'food' },
  'hotels': { name: 'Hotels & Accommodation', icon: '\uD83C\uDFE8', filter: 'hotel' },
  'retail': { name: 'Retail & E-commerce', icon: '\uD83D\uDECD', filter: 'retail' },
  'medical': { name: 'Medical & Health', icon: '\uD83C\uDFE5', filter: 'medical' },
  'tourism': { name: 'Tourism & Travel', icon: '\u2708', filter: 'tourism' },
  'professionals': { name: 'Professional Services', icon: '\uD83D\uDCBC', filter: 'prof' },
  'beauty': { name: 'Beauty & Wellness', icon: '\uD83D\uDC87', filter: 'beauty' },
  'automotive': { name: 'Automotive', icon: '\uD83D\uDE97', filter: 'auto' },
  'no-website': { name: 'ALL \u2014 No Website', icon: '\uD83D\uDCE1', filter: 'none' },
  'easy-reach': { name: 'Easiest to Reach', icon: '\uD83D\uDCDE', filter: 'easy' },
};

fs.writeFileSync('crm-data.js', `const CRM = ${JSON.stringify(crmData)};
const CRM_CATS = ${JSON.stringify(sortedCats)};
const CRM_SERVICES = ${JSON.stringify(svcPrices)};
const CRM_LIST_META = ${JSON.stringify(listMeta)};
const CRM_LIST_DATA = {
  'hot-leads': ${JSON.stringify(count('hot-leads', d => d.needsWebsite && d.phone.length > 0))},
  'restaurants': ${JSON.stringify(count('restaurants', d => foodCats.includes(d.cat) && d.needsWebsite))},
  'hotels': ${JSON.stringify(count('hotels', d => hotelCats.includes(d.cat)))},
  'retail': ${JSON.stringify(count('retail', d => retailCats.includes(d.cat)))},
  'medical': ${JSON.stringify(count('medical', d => medicalCats.includes(d.cat)))},
  'tourism': ${JSON.stringify(count('tourism', d => tourismCats.includes(d.cat)))},
  'professionals': ${JSON.stringify(count('professionals', d => profCats.includes(d.cat)))},
  'beauty': ${JSON.stringify(count('beauty', d => beautyCats.includes(d.cat)))},
  'automotive': ${JSON.stringify(count('automotive', d => autoCats.includes(d.cat)))},
  'no-website': ${JSON.stringify(count('no-website', d => d.needsWebsite))},
  'easy-reach': ${JSON.stringify(count('easy-reach', d => d.needsWebsite && d.email.length > 0 && d.phone.length > 0))},
};
const CRM_FOOD = ${JSON.stringify(foodCats)};
const CRM_HOTEL = ${JSON.stringify(hotelCats)};
const CRM_RETAIL = ${JSON.stringify(retailCats)};
const CRM_MEDICAL = ${JSON.stringify(medicalCats)};
const CRM_TOURISM = ${JSON.stringify(tourismCats)};
const CRM_PROF = ${JSON.stringify(profCats)};
const CRM_BEAUTY = ${JSON.stringify(beautyCats)};
const CRM_AUTO = ${JSON.stringify(autoCats)};`);

const result = {};
for (const [k, v] of Object.entries(listMeta)) {
  let fn;
  if (v.filter === 'hot') fn = d => d.needsWebsite && d.phone.length > 0;
  else if (v.filter === 'food') fn = d => foodCats.includes(d.cat) && d.needsWebsite;
  else if (v.filter === 'hotel') fn = d => hotelCats.includes(d.cat);
  else if (v.filter === 'retail') fn = d => retailCats.includes(d.cat);
  else if (v.filter === 'medical') fn = d => medicalCats.includes(d.cat);
  else if (v.filter === 'tourism') fn = d => tourismCats.includes(d.cat);
  else if (v.filter === 'prof') fn = d => profCats.includes(d.cat);
  else if (v.filter === 'beauty') fn = d => beautyCats.includes(d.cat);
  else if (v.filter === 'auto') fn = d => autoCats.includes(d.cat);
  else if (v.filter === 'none') fn = d => d.needsWebsite;
  else if (v.filter === 'easy') fn = d => d.needsWebsite && d.email.length > 0 && d.phone.length > 0;
  const items = crmData.filter(fn);
  result[k] = { count: items.length, totalValue: items.reduce((s, d) => s + d.totalValue, 0) };
}
console.log(`CRM built: ${crmData.length} businesses, ${sortedCats.length} categories`);
for (const [k, v] of Object.entries(listMeta)) {
  console.log(`  ${v.icon} ${v.name}: ${result[k].count} leads, $${result[k].totalValue.toLocaleString()}`);
}
