const axios = require('axios');
const fs = require('fs');
const config = require('./config.json');

const API_KEY = config.apiKey;
const SEARCH_DELAY = config.searchDelayMs || 200;
const BATCH_SIZE = config.batchSize || 50;
const MAX_RESULTS = config.maxResults || 3;

if (!API_KEY || API_KEY === 'YOUR_GOOGLE_API_KEY_HERE') {
    console.log('⚠  Please set your Google API key in config.json');
    console.log('   Get one at: https://console.cloud.google.com/');
    process.exit(1);
}

// ─── Load businesses ───
const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));

// ─── Load or init verification progress ───
let verified = {};
if (fs.existsSync('verify-progress.json')) {
    verified = JSON.parse(fs.readFileSync('verify-progress.json', 'utf8'));
    console.log('Resuming from previous verification...');
}

// ─── Google Places API Text Search ───
async function searchPlaces(businessName, category) {
    const query = `${businessName} ${category || ''} Kigali Rwanda`;
    const url = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(query)}&key=${API_KEY}`;
    
    try {
        const res = await axios.get(url, { timeout: 10000 });
        if (res.data.status === 'OK' && res.data.results.length > 0) {
            return res.data.results.slice(0, MAX_RESULTS).map(p => ({
                name: p.name,
                address: p.formatted_address || '',
                phone: '',  // need place details for phone
                website: '', // need place details for website
                rating: p.rating || 0,
                placeId: p.place_id,
                types: p.types || []
            }));
        }
        return null;
    } catch (e) {
        return null;
    }
}

// ─── Google Places Detail (get phone + website) ───
async function getPlaceDetails(placeId) {
    const url = `https://maps.googleapis.com/maps/api/place/details/json?place_id=${placeId}&fields=name,formatted_phone_number,website,formatted_address,rating,types&key=${API_KEY}`;
    
    try {
        const res = await axios.get(url, { timeout: 10000 });
        if (res.data.status === 'OK') {
            const d = res.data.result;
            return {
                phone: d.formatted_phone_number || '',
                website: d.website || '',
                address: d.formatted_address || '',
                rating: d.rating || 0
            };
        }
        return null;
    } catch (e) {
        return null;
    }
}

// ─── Google Custom Search (fallback for website) ───
async function searchWeb(businessName) {
    const query = `${businessName} Kigali Rwanda official website`;
    const url = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(query)}&key=${API_KEY}&cx=${config.searchEngineId || ''}`;
    
    try {
        const res = await axios.get(url, { timeout: 10000 });
        if (res.data.items && res.data.items.length > 0) {
            const result = res.data.items[0];
            return {
                website: result.link || '',
                title: result.title || '',
                snippet: result.snippet || ''
            };
        }
        return null;
    } catch (e) {
        return null;
    }
}

// ─── Verify a single business ───
async function verifyBusiness(business) {
    const name = business.name;
    const category = business.category || '';
    
    // 1. Try Google Places
    const places = await searchPlaces(name, category);
    if (places && places.length > 0) {
        // Get details for the best match
        const details = await getPlaceDetails(places[0].placeId);
        if (details) {
            return {
                verified: true,
                source: 'google_places',
                website: details.website || business.website || '',
                phone: details.phone || business.phone || business.mobile || '',
                address: details.address || business.address || '',
                rating: details.rating,
                placeName: places[0].name,
                matchConfidence: 'high'
            };
        }
    }
    
    // 2. Try Custom Search (if configured)
    if (config.customSearchEnabled && config.searchEngineId) {
        const webResult = await searchWeb(name);
        if (webResult && webResult.website) {
            return {
                verified: true,
                source: 'google_search',
                website: webResult.website,
                title: webResult.title,
                snippet: webResult.snippet,
                phone: business.phone || business.mobile || '',
                address: business.address || '',
                matchConfidence: 'medium'
            };
        }
    }
    
    return {
        verified: false,
        source: 'none',
        website: business.website || '',
        phone: business.phone || business.mobile || '',
        address: business.address || '',
        matchConfidence: 'low'
    };
}

// ─── Main loop ───
async function main() {
    console.log('Google Maps Business Verification');
    console.log('Total businesses: ' + businesses.length);
    console.log('Already verified: ' + Object.keys(verified).length);
    console.log('');
    
    let found = 0;
    let notFound = 0;
    let websitesUpdated = 0;
    
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        const id = i + 1;
        
        // Skip if already verified in this session
        if (verified[id]) continue;
        
        // Only verify businesses that need it
        const currentWebsite = b.verifiedWebsite || b.website || '';
        const needsVerification = !currentWebsite || b.websiteWorking !== true || !b.googleVerified;
        
        if (!needsVerification) {
            // Mark as already good
            verified[id] = { status: 'skipped', website: currentWebsite };
            continue;
        }
        
        process.stdout.write(`[${id}/${businesses.length}] ${b.name.substring(0, 50).padEnd(50)} `);
        
        const result = await verifyBusiness(b);
        
        if (result.verified && result.website) {
            found++;
            if (result.website !== currentWebsite) {
                b.verifiedWebsite = result.website;
                b.websiteWorking = true;
                b.googleVerified = true;
                b.googlePlaceName = result.placeName || '';
                b.googleRating = result.rating || 0;
                websitesUpdated++;
                process.stdout.write('✓ ' + result.website.substring(0, 50) + '\n');
            } else {
                b.googleVerified = true;
                process.stdout.write('✓ (same)\n');
            }
            
            // Update phone/address if Google has better data
            if (result.phone && !b.phone && !b.mobile) {
                b.phone = result.phone;
            }
            if (result.address && !b.address) {
                b.address = result.address;
            }
            
            verified[id] = { status: 'found', website: result.website, source: result.source };
        } else {
            notFound++;
            b.googleVerified = false;
            verified[id] = { status: 'not_found' };
            process.stdout.write('✗ not found\n');
        }
        
        // Save progress periodically
        if ((i + 1) % BATCH_SIZE === 0) {
            fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
            fs.writeFileSync('verify-progress.json', JSON.stringify(verified, null, 2));
            console.log(`  → Progress saved. Found: ${found} | Not found: ${notFound} | Updated: ${websitesUpdated}\n`);
        }
        
        await new Promise(r => setTimeout(r, SEARCH_DELAY));
    }
    
    // Final save
    fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
    fs.writeFileSync('verify-progress.json', JSON.stringify(verified, null, 2));
    
    console.log('\n═══════════════════════════════════');
    console.log('VERIFICATION COMPLETE');
    console.log('Total businesses: ' + businesses.length);
    console.log('Websites found:    ' + found);
    console.log('Not found:         ' + notFound);
    console.log('Websites updated:  ' + websitesUpdated);
    console.log('═══════════════════════════════════\n');
    
    // Generate verified Excel
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Verified Directory');
    
    ws.columns = [
        { header: '#', key: 'id', width: 4 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Email', key: 'email', width: 35 },
        { header: 'Phone', key: 'phone', width: 18 },
        { header: 'Website', key: 'website', width: 42 },
        { header: 'Website Status', key: 'ws', width: 14 },
        { header: 'Google Rating', key: 'rating', width: 10 },
        { header: 'Address', key: 'address', width: 45 },
        { header: 'Verified Source', key: 'source', width: 18 }
    ];
    
    const hdrStyle = { font: { bold: true, color: { argb: 'FFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1565C0' } } };
    ws.getRow(1).eachCell(c => c.style = hdrStyle);
    
    let rowId = 1;
    for (const b of businesses) {
        const web = b.verifiedWebsite || b.website || '';
        const source = verified[rowId]?.source || (b.googleVerified ? 'google' : 'original');
        ws.addRow({
            id: rowId,
            name: b.name,
            cat: b.category || '',
            email: b.email || '',
            phone: b.phone || b.mobile || '',
            website: web,
            ws: web ? (b.websiteWorking === true ? 'Working' : 'Unknown') : 'No Website',
            rating: b.googleRating || '',
            address: b.address || '',
            source: source
        });
        rowId++;
    }
    
    await wb.xlsx.writeFile('Kigali_Verified.xlsx');
    console.log('Saved: Kigali_Verified.xlsx');
}

main().catch(e => {
    console.error('Error:', e.message);
    // Save on crash
    fs.writeFileSync('verify-progress.json', JSON.stringify(verified, null, 2));
});
