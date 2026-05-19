const axios = require('axios');
const fs = require('fs');

const API_KEY = 'AIzaSyBWsc2Pi6N-fyBwjT-dtqg8rSgTLYMJmfk';
const DELAY = 600;

const SEARCHES = [
  ['Restaurants', 'restaurants in Kigali Rwanda'],
  ['Cafes', 'cafes in Kigali Rwanda'],
  ['Hotels', 'hotels in Kigali Rwanda'],
  ['Shops', 'shops in Kigali Rwanda'],
  ['Clinics', 'clinics in Kigali Rwanda'],
  ['Pharmacies', 'pharmacies in Kigali Rwanda'],
  ['Beauty Salons', 'beauty salons in Kigali Rwanda'],
  ['Schools', 'schools in Kigali Rwanda'],
  ['Banks', 'banks in Kigali Rwanda'],
  ['Lawyers', 'lawyers in Kigali Rwanda'],
  ['Supermarkets', 'supermarkets in Kigali Rwanda'],
  ['Car Repair', 'car repair in Kigali Rwanda'],
  ['Gyms', 'gyms in Kigali Rwanda'],
  ['Barbershops', 'barbershops in Kigali Rwanda'],
  ['Real Estate', 'real estate agents in Kigali Rwanda'],
  ['Hardware Stores', 'hardware stores in Kigali Rwanda'],
  ['Fashion Stores', 'fashion stores in Kigali Rwanda'],
  ['Electronics', 'electronics stores in Kigali Rwanda'],
  ['Furniture', 'furniture stores in Kigali Rwanda'],
  ['Bakeries', 'bakeries in Kigali Rwanda'],
  ['Bars', 'bars in Kigali Rwanda'],
  ['Dentists', 'dentists in Kigali Rwanda'],
  ['Opticians', 'opticians in Kigali Rwanda'],
  ['Travel Agents', 'travel agents in Kigali Rwanda'],
  ['Car Rental', 'car rental in Kigali Rwanda'],
  ['Laundry', 'laundry services in Kigali Rwanda'],
  ['Printing', 'printing services in Kigali Rwanda'],
  ['Catering', 'catering in Kigali Rwanda'],
  ['Logistics', 'logistics companies in Kigali Rwanda'],
  ['Security', 'security services in Kigali Rwanda'],
  ['Construction', 'construction companies in Kigali Rwanda'],
  ['IT Services', 'IT services in Kigali Rwanda'],
  ['Accountants', 'accountants in Kigali Rwanda'],
  ['Architects', 'architects in Kigali Rwanda'],
  ['Photographers', 'photographers in Kigali Rwanda'],
  ['Event Venues', 'event venues in Kigali Rwanda'],
  ['Guest Houses', 'guest houses in Kigali Rwanda'],
  ['Spa', 'spa in Kigali Rwanda'],
  ['Car Wash', 'car wash in Kigali Rwanda'],
  ['Delivery Services', 'delivery services in Kigali Rwanda'],
  ['Pharmacies', 'pharmaceutical companies in Kigali Rwanda'],
  ['Transport', 'transport services in Kigali Rwanda'],
  ['Courier', 'courier services in Kigali Rwanda'],
  ['Bookstores', 'bookstores in Kigali Rwanda'],
];

const CAT_MAP = {
  restaurant: 'Restaurants',
  cafe: 'Cafes',
  fast_food: 'Restaurants',
  meal_delivery: 'Restaurants',
  meal_takeaway: 'Restaurants',
  hotel: 'Hotels',
  lodging: 'Hotels',
  guest_house: 'Bed and Breakfast',
  campground: 'Bed and Breakfast',
  supermarket: 'Supermarkets',
  convenience_store: 'Shops',
  grocery: 'Food Retailers',
  liquor_store: 'Shops',
  food: 'Food Retailers',
  store: 'Shops',
  clothing_store: 'Fashion',
  boutique: 'Fashion',
  shoe_store: 'Fashion',
  jewelry_store: 'Fashion',
  doctor: 'Doctors and Clinics',
  clinic: 'Doctors and Clinics',
  hospital: 'Doctors and Clinics',
  health: 'Doctors and Clinics',
  pharmacy: 'Pharmacies',
  beauty_salon: 'Beauty Professionals',
  hair_care: 'Beauty Professionals',
  barber: 'Beauty Professionals',
  spa: 'Beauty Professionals',
  lawyer: 'Lawyers',
  real_estate_agency: 'RealEstate Agents',
  school: 'Education',
  university: 'Education',
  secondary_school: 'Education',
  primary_school: 'Education',
  car_repair: 'Automotive',
  auto_parts: 'Automotive',
  gym: 'Fitness',
  fitness: 'Fitness',
  bar: 'Pubs and Clubs',
  night_club: 'Pubs and Clubs',
  bakery: 'Food Retailers',
  dentist: 'Doctors and Clinics',
  travel_agency: 'Travel Agents',
  car_rental: 'Car Rental',
  hardware_store: 'Hardware Stores',
  accountant: 'Audit and Accounting',
  architect: 'Architectural Services',
  electrician: 'Construction',
  plumber: 'Construction',
  printing: 'Printing',
  caterer: 'Catering',
  moving_company: 'Logistics',
  storage: 'Logistics',
  security: 'Security Services',
  insurance_agency: 'Insurance',
  bank: 'Banks, Credit Unions',
  furniture_store: 'Furniture',
  electronics_store: 'Electronic Equipment',
  bookstore: 'Books',
  laundry: 'Dry Cleaning',
  dry_cleaning: 'Dry Cleaning',
  photographer: 'Photography',
  it: 'Information Technology',
  training: 'Training',
  consultant: 'Consultants',
  car_wash: 'Vehicle Services',
  courier: 'Logistics',
  delivery: 'Logistics',
  general_contractor: 'Construction',
};

function mapCat(types) {
  if (!types) return 'General Business';
  for (const t of types) {
    const clean = t.replace(/_/g, '_');
    if (CAT_MAP[t]) return CAT_MAP[t];
  }
  return 'General Business';
}

function normalizePhone(p) {
  if (!p) return '';
  let s = p.replace(/[^0-9+]/g, '');
  if (s.startsWith('00250')) s = '+250' + s.slice(5);
  else if (s.startsWith('250') && !s.startsWith('+')) s = '+250' + s.slice(3);
  else if (s.startsWith('0')) s = '+250' + s.slice(1);
  else if (!s.startsWith('+')) s = '+' + s;
  return s.length >= 10 ? s.substring(0, 13) : '';
}

function getIndustryNeeds(cat) {
  const map = {
    'Restaurants': ['Website','Online Ordering','Booking System','POS System','Inventory Mgmt','Marketing Automation'],
    'Hotels': ['Website','Booking System','POS System','CRM','Marketing Automation','Communication'],
    'Bed and Breakfast': ['Website','Booking System','Marketing Automation','Social Media Mgmt','Communication','Payment Gateway'],
    'Cafes': ['Website','Online Ordering','POS System','Inventory Mgmt','Marketing Automation','Social Media Mgmt'],
    'Food Retailers': ['Website','Online Ordering','POS System','Inventory Mgmt','Delivery Logistics','Marketing Automation'],
    'Shops': ['Website','E-commerce','POS System','Inventory Mgmt','Accounting/ERP','Marketing Automation'],
    'Fashion': ['Website','E-commerce','Social Media Mgmt','Inventory Mgmt','Marketing Automation','POS System'],
    'Supermarkets': ['Website','E-commerce','Online Ordering','POS System','Inventory Mgmt','Delivery Logistics'],
    'Doctors and Clinics': ['Website','Patient Mgmt','Booking System','CRM','SMS/Notification','Cloud Backup'],
    'Pharmacies': ['Website','Inventory Mgmt','POS System','SMS/Notification','Accounting/ERP','Online Ordering'],
    'Beauty Professionals': ['Website','Booking System','Social Media Mgmt','Marketing Automation','POS System','SMS/Notification'],
    'Automotive': ['Website','Inventory Mgmt','POS System','CRM','SMS/Notification','Accounting/ERP'],
    'Vehicle Services': ['Website','Booking System','POS System','CRM','Marketing Automation','SMS/Notification'],
    'Education': ['Website','Learning Mgmt','SMS/Notification','CRM','Cloud Backup','Communication'],
    'RealEstate Agents': ['Website','CRM','Marketing Automation','Social Media Mgmt','Communication','Analytics & Reporting'],
    'Lawyers': ['Website','CRM','Cloud Backup','Communication','Accounting/ERP','Booking System'],
    'Construction': ['Website','CRM','Accounting/ERP','Cloud Backup','Analytics & Reporting','Communication'],
    'Travel Agents': ['Website','Booking System','CRM','Marketing Automation','Communication','Payment Gateway'],
    'Car Rental': ['Website','Booking System','CRM','Marketing Automation','Communication','Payment Gateway'],
    'Catering': ['Website','Online Ordering','Delivery Logistics','Inventory Mgmt','Marketing Automation','Communication'],
    'Dry Cleaning': ['Website','Online Ordering','POS System','Inventory Mgmt','SMS/Notification','Marketing Automation'],
    'Logistics': ['Website','CRM','Delivery Logistics','Inventory Mgmt','Analytics & Reporting','Communication'],
    'Transport Services': ['Website','CRM','Booking System','Communication','Analytics & Reporting','Marketing Automation'],
    'Hardware Stores': ['Website','E-commerce','Inventory Mgmt','POS System','Accounting/ERP','Delivery Logistics'],
    'Printing': ['Website','E-commerce','Inventory Mgmt','POS System','Analytics & Reporting','Marketing Automation'],
    'Electronic Equipment': ['Website','E-commerce','Inventory Mgmt','POS System','Accounting/ERP','Marketing Automation'],
    'Furniture': ['Website','E-commerce','Inventory Mgmt','POS System','Delivery Logistics','Marketing Automation'],
    'Security Services': ['Website','CRM','Security Systems','Communication','Analytics & Reporting','Cloud Backup'],
    'Consultants': ['Website','CRM','Marketing Automation','Communication','Cloud Backup','Accounting/ERP'],
    'Audit and Accounting': ['Website','CRM','Marketing Automation','Cloud Backup','Communication','Accounting/ERP'],
    'Pubs and Clubs': ['Website','Booking System','POS System','Social Media Mgmt','Marketing Automation','SMS/Notification'],
    'Fitness': ['Website','Booking System','CRM','Social Media Mgmt','Marketing Automation','POS System'],
    'Books': ['Website','E-commerce','Inventory Mgmt','POS System','Accounting/ERP','Marketing Automation'],
    'Home and Garden': ['Website','E-commerce','Inventory Mgmt','POS System','Delivery Logistics','Marketing Automation'],
    'Banks, Credit Unions': ['Website','CRM','Communication','Security Systems','Analytics & Reporting','SMS/Notification'],
    'Information Technology': ['Website','CRM','Marketing Automation','Cloud Backup','Communication','Security Systems'],
    'Photography': ['Website','Booking System','Social Media Mgmt','Marketing Automation','Cloud Backup','Communication'],
    'Architectural Services': ['Website','CRM','Cloud Backup','Communication','Analytics & Reporting','Accounting/ERP'],
    'Insurance': ['Website','CRM','Marketing Automation','Communication','Cloud Backup','Accounting/ERP'],
    'Non-Profit': ['Website','CRM','Marketing Automation','Communication','Cloud Backup','Analytics & Reporting'],
    'Event Services': ['Website','Booking System','CRM','Social Media Mgmt','Marketing Automation','Communication'],
    'Opticians': ['Website','Inventory Mgmt','POS System','CRM','Booking System','SMS/Notification'],
    'General Business': ['Website','CRM','Marketing Automation','Cloud Backup','Communication','Social Media Mgmt'],
  };
  return map[cat] || ['Website','CRM','Marketing Automation','Cloud Backup','Communication','Social Media Mgmt'];
}

function getSvcPrices() {
  return {
    'Website':350,'E-commerce':800,'Booking System':600,'POS System':500,'CRM':400,'Inventory Mgmt':450,
    'Accounting/ERP':700,'Marketing Automation':350,'Communication':300,'Online Ordering':500,'HR & Payroll':400,
    'Analytics & Reporting':300,'Security Systems':500,'Cloud Backup':200,'Social Media Mgmt':250,'SMS/Notification':200,
    'Payment Gateway':300,'Delivery Logistics':400,'Learning Mgmt':500,'Patient Mgmt':600
  };
}

const seen = new Set();
const businesses = [];

async function searchText(query, pageToken) {
  try {
    const params = { query, key: API_KEY, language: 'en' };
    if (pageToken) params.pagetoken = pageToken;
    const res = await axios.get('https://maps.googleapis.com/maps/api/place/textsearch/json', {
      params, timeout: 15000,
    });
    if (res.data.status === 'OK' || res.data.status === 'ZERO_RESULTS') return res.data;
    console.error(`  API error: ${res.data.status} - ${res.data.error_message || ''}`);
    return null;
  } catch (e) {
    console.error(`  Request failed: ${e.message}`);
    return null;
  }
}

async function main() {
  console.log('=== Google Places API Scanner for Kigali ===\n');

  for (let si = 0; si < SEARCHES.length; si++) {
    const [label, query] = SEARCHES[si];
    console.log(`[${si + 1}/${SEARCHES.length}] ${label}...`);
    let token = null;
    for (let pg = 0; pg < 3; pg++) {
      const data = await searchText(query, token);
      if (!data || !data.results) break;
      for (const p of data.results) {
        if (seen.has(p.place_id)) continue;
        seen.add(p.place_id);
        const name = p.name || '';
        if (!name) continue;
        const phone = normalizePhone(p.international_phone_number || p.formatted_phone_number || '');
        const website = p.website || '';
        const address = p.formatted_address || '';
        const types = p.types || [];
        const cat = mapCat(types);
        const services = getIndustryNeeds(cat);
        const svcPrices = getSvcPrices();
        const totalValue = services.reduce((s, n) => s + (svcPrices[n] || 300), 0);
        const ws = website ? 'working' : 'none';

        businesses.push({
          place_id: p.place_id,
          name, phone, website, ws, address, cat,
          email: '',
          services: services.join(', '),
          totalValue,
          needsWebsite: !website,
          priority: !website ? 'hot' : 'low',
          source: 'Google Places',
        });
      }
      console.log(`  Page ${pg + 1}: ${data.results.length} (${seen.size} unique)`);
      token = data.next_page_token || null;
      if (!token) break;
      await new Promise(r => setTimeout(r, 1500));
    }
    await new Promise(r => setTimeout(r, DELAY));
  }

  // Deduplicate
  const nameSet = new Set();
  const unique = [];
  for (const b of businesses) {
    const key = b.name.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!nameSet.has(key)) { nameSet.add(key); unique.push(b); }
  }

  console.log(`\n=== Results ===`);
  console.log(`Total: ${unique.length}`);
  console.log(`With website: ${unique.filter(b => b.website).length}`);
  console.log(`With phone: ${unique.filter(b => b.phone).length}`);
  console.log(`Hot leads: ${unique.filter(b => b.needsWebsite && b.phone).length}`);

  fs.writeFileSync('places_raw.json', JSON.stringify(unique, null, 2));
  buildCRM(unique);
}

function buildCRM(data) {
  const svcPrices = getSvcPrices();
  const crmData = data.map((b, i) => ({
    id: i + 1, name: b.name, cat: b.cat, email: b.email,
    phone: b.phone, website: b.website, ws: b.ws,
    address: b.address.substring(0, 80),
    services: b.services, totalValue: b.totalValue,
    needsWebsite: b.needsWebsite, priority: b.priority, source: b.source,
  }));

  const cats = {};
  crmData.forEach(d => { cats[d.cat] = (cats[d.cat] || 0) + 1; });
  const sortedCats = Object.entries(cats).sort((a, b) => b[1] - a[1]);

  const foodCats = ['Restaurants','Food Retailers','Catering','Cafes','Supermarkets'];
  const hotelCats = ['Hotels','Bed and Breakfast'];
  const retailCats = ['Fashion','Shops','Hardware Stores','Electronic Equipment','Furniture','Supermarkets','Books','Home and Garden'];
  const medicalCats = ['Doctors and Clinics','Pharmacies','Opticians'];
  const tourismCats = ['Travel Agents','Car Rental','Transport Services'];
  const profCats = ['Lawyers','Consultants','Audit and Accounting','Architectural Services','Insurance','Information Technology'];
  const beautyCats = ['Beauty Professionals','Fitness'];
  const autoCats = ['Automotive','Vehicle Services','Car Rental'];

  function count(key, fn) {
    const items = crmData.filter(fn);
    return { count: items.length, totalValue: items.reduce((s, d) => s + d.totalValue, 0) };
  }

  const listMeta = {
    'hot-leads': { name: 'Hot Leads — No Website', icon: '🔥', filter: 'hot' },
    'restaurants': { name: 'Restaurants & Food', icon: '🍽', filter: 'food' },
    'hotels': { name: 'Hotels & Accommodation', icon: '🏨', filter: 'hotel' },
    'retail': { name: 'Retail & E-commerce', icon: '🛍', filter: 'retail' },
    'medical': { name: 'Medical & Health', icon: '🏥', filter: 'medical' },
    'tourism': { name: 'Tourism & Travel', icon: '✈', filter: 'tourism' },
    'professionals': { name: 'Professional Services', icon: '💼', filter: 'prof' },
    'beauty': { name: 'Beauty & Wellness', icon: '💇', filter: 'beauty' },
    'automotive': { name: 'Automotive', icon: '🚗', filter: 'auto' },
    'no-website': { name: 'ALL — No Website', icon: '📡', filter: 'none' },
    'easy-reach': { name: 'Easiest to Reach', icon: '📞', filter: 'easy' },
  };

  const content = `const CRM = ${JSON.stringify(crmData)};
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
const CRM_AUTO = ${JSON.stringify(autoCats)};`;

  fs.writeFileSync('crm-data.js', content);
  console.log(`\nCRM built: ${crmData.length} businesses`);
}

main().catch(console.error);
