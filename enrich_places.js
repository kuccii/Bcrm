const axios = require('axios');
const fs = require('fs');

const API_KEY = 'AIzaSyBWsc2Pi6N-fyBwjT-dtqg8rSgTLYMJmfk';
const raw = JSON.parse(fs.readFileSync('places_raw.json', 'utf8'));
const DELAY = 60;

function normalizePhone(p) {
  if (!p) return '';
  let s = p.replace(/[^0-9+]/g, '');
  if (s.startsWith('00250')) s = '+250' + s.slice(5);
  else if (s.startsWith('250') && !s.startsWith('+')) s = '+250' + s.slice(3);
  else if (s.startsWith('0')) s = '+250' + s.slice(1);
  else if (!s.startsWith('+')) s = '+' + s;
  return s.length >= 10 ? s.substring(0, 13) : '';
}

async function getDetails(placeId) {
  try {
    const res = await axios.get('https://maps.googleapis.com/maps/api/place/details/json', {
      params: {
        place_id: placeId,
        fields: 'place_id,name,formatted_phone_number,international_phone_number,website',
        key: API_KEY,
      },
      timeout: 10000,
    });
    if (res.data.status === 'OK') return res.data.result;
    return null;
  } catch { return null; }
}

async function main() {
  const withSite = raw.filter(b => b.website).length;
  const withPhone = raw.filter(b => b.phone).length;
  console.log(`Starting enrichment: ${raw.length} businesses (${withSite} have website, ${withPhone} have phone)`);

  let enriched = 0;
  for (let i = 0; i < raw.length; i++) {
    const b = raw[i];
    if (!b.place_id) continue;
    if (b.phone || b.website) continue;

    const d = await getDetails(b.place_id);
    if (d) {
      b.phone = normalizePhone(d.international_phone_number || d.formatted_phone_number || '');
      b.website = d.website || '';
      b.ws = b.website ? 'working' : 'none';
      b.needsWebsite = !b.website;
      b.priority = !b.website ? 'hot' : 'low';
      enriched++;
    }

    if ((i + 1) % 100 === 0) {
      const w = raw.filter(x => x.website).length;
      const p = raw.filter(x => x.phone).length;
      console.log(`  ${i + 1}/${raw.length} (enriched: ${enriched}, website: ${w}, phone: ${p})`);
      fs.writeFileSync('places_raw.json', JSON.stringify(raw, null, 2));
    }
    await new Promise(r => setTimeout(r, DELAY));
  }

  const ws = raw.filter(b => b.website).length;
  const ph = raw.filter(b => b.phone).length;
  const hot = raw.filter(b => b.needsWebsite && b.phone).length;
  console.log(`\nDone. Website: ${ws}, Phone: ${ph}, Hot leads: ${hot}`);
  fs.writeFileSync('places_raw.json', JSON.stringify(raw, null, 2));
  buildCRM(raw);
}

function getSvcPrices() {
  return {
    'Website':350,'E-commerce':800,'Booking System':600,'POS System':500,'CRM':400,'Inventory Mgmt':450,
    'Accounting/ERP':700,'Marketing Automation':350,'Communication':300,'Online Ordering':500,'HR & Payroll':400,
    'Analytics & Reporting':300,'Security Systems':500,'Cloud Backup':200,'Social Media Mgmt':250,'SMS/Notification':200,
    'Payment Gateway':300,'Delivery Logistics':400,'Learning Mgmt':500,'Patient Mgmt':600
  };
}

function buildCRM(data) {
  const svcPrices = getSvcPrices();
  const crmData = data.map((b, i) => ({
    id: i + 1, name: b.name, cat: b.cat, email: b.email || '',
    phone: b.phone || '', website: b.website || '', ws: b.ws || 'none',
    address: (b.address || '').substring(0, 80),
    services: b.services || 'Website, CRM, Marketing Automation, Cloud Backup, Communication, Social Media Mgmt',
    totalValue: b.totalValue || 1850,
    needsWebsite: !b.website,
    priority: b.priority || 'hot',
    source: b.source || 'Google Places',
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
  console.log(`\n=== CRM Saved: ${crmData.length} businesses ===`);
  const easyReach = crmData.filter(d => d.needsWebsite && d.email && d.phone).length;
  console.log(`Hot leads: ${crmData.filter(d => d.needsWebsite && d.phone).length}`);
  console.log(`Easiest to reach: ${easyReach}`);
}

main().catch(console.error);
