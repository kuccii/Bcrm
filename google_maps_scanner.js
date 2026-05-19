const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const fs = require('fs');

const DELAY_MS = 1000;

async function searchGoogleMaps(businessName) {
    try {
        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(businessName + " Rwanda")}&hl=en`;
        const response = await axios.get(searchUrl, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
            },
            timeout: 15000
        });
        
        const $ = cheerio.load(response.data);
        let website = '';
        let mapsLink = '';
        let foundInfo = '';
        
        $('a[href*="maps.google"]').first().each((i, el) => {
            mapsLink = $(el).attr('href') || '';
        });
        
        $('a[href^="http"]:not([href*="google"]):not([href*="youtube"]):not([href*="facebook"]):not([href*="twitter"]):not([href*="linkedin"]):not([href*="instagram"])').each((i, el) => {
            const href = $(el).attr('href');
            const text = $(el).text().toLowerCase();
            if (!website && (text.includes('website') || text.includes('.com') || text.includes('.rw') || text.includes('.org'))) {
                website = href.split('&')[0];
            }
        });
        
        const description = $('div[data-sncf]').first().text().trim() || 
                           $('span[style*="font-size"]').first().text().trim();
        
        return { website, mapsLink, description };
    } catch (e) {
        return { website: '', mapsLink: '', description: '', error: e.message };
    }
}

async function checkWebsiteWorking(url) {
    if (!url || !url.startsWith('http')) return false;
    try {
        const response = await axios.get(url, {
            timeout: 10000,
            validateStatus: (status) => status < 500
        });
        return response.status >= 200 && response.status < 400;
    } catch (e) {
        return false;
    }
}

async function loadProgress() {
    if (fs.existsSync('progress.json')) {
        return JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    }
    return [];
}

async function saveProgress(data) {
    fs.writeFileSync('progress.json', JSON.stringify(data, null, 2));
}

async function createExcel(businesses, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Businesses');
    
    worksheet.columns = [
        { header: 'Business Name', key: 'name', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Address', key: 'address', width: 35 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone', key: 'phone', width: 15 },
        { header: 'Mobile', key: 'mobile', width: 15 },
        { header: 'Contact Person', key: 'contactPerson', width: 25 },
        { header: 'Website', key: 'website', width: 35 },
        { header: 'Website Source', key: 'websiteSource', width: 15 },
        { header: 'Website Working', key: 'websiteWorking', width: 15 },
        { header: 'Google Maps', key: 'mapsLink', width: 40 },
        { header: 'Profile URL', key: 'url', width: 40 }
    ];
    
    for (const b of businesses) {
        worksheet.addRow({
            name: b.name,
            category: b.category || 'N/A',
            address: b.address || 'N/A',
            email: b.email || 'N/A',
            phone: b.phone || 'N/A',
            mobile: b.mobile || 'N/A',
            contactPerson: b.contactPerson || 'N/A',
            website: b.googleWebsite || b.website || 'N/A',
            websiteSource: b.googleWebsite ? 'Google' : (b.website ? 'RwandaYP' : 'N/A'),
            websiteWorking: b.googleWebsiteWorking !== undefined ? (b.googleWebsiteWorking ? 'Yes' : 'No') : 
                           (b.websiteWorking === true ? 'Yes' : (b.websiteWorking === false ? 'No' : 'N/A')),
            mapsLink: b.googleMapsLink || 'N/A',
            url: b.url || ''
        });
    }
    
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Excel saved: ${outputPath}`);
}

async function main() {
    console.log('Loading existing data...');
    let businesses = await loadProgress();
    console.log(`Loaded ${businesses.length} businesses`);
    
    const withoutWebsite = businesses.filter(b => !b.website && !b.googleWebsite);
    console.log(`Businesses without website: ${withoutWebsite.length}`);
    
    console.log('\nSearching Google Maps for businesses without websites...\n');
    
    let processed = 0;
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        
        if (!b.website && !b.googleWebsite) {
            console.log(`[${i + 1}/${businesses.length}] Searching: ${b.name}`);
            
            const result = await searchGoogleMaps(b.name);
            
            if (result.website) {
                b.googleWebsite = result.website;
                b.googleMapsLink = result.mapsLink;
                console.log(`  Found website: ${result.website}`);
                
                b.googleWebsiteWorking = await checkWebsiteWorking(result.website);
                console.log(`  Website working: ${b.googleWebsiteWorking ? 'Yes' : 'No'}`);
            } else if (result.mapsLink) {
                b.googleMapsLink = result.mapsLink;
                console.log(`  Found Google Maps link`);
            }
            
            if ((i + 1) % 10 === 0) {
                await saveProgress(businesses);
                console.log(`  Progress saved`);
            }
            
            await new Promise(resolve => setTimeout(resolve, DELAY_MS));
        }
        processed++;
    }
    
    await saveProgress(businesses);
    
    const hasWebsite = businesses.filter(b => b.website || b.googleWebsite).length;
    const websiteWorks = businesses.filter(b => b.websiteWorking === true || b.googleWebsiteWorking === true).length;
    
    console.log('\n=== Final Results ===');
    console.log(`Total: ${businesses.length}`);
    console.log(`Has website: ${hasWebsite}`);
    console.log(`Website working: ${websiteWorks}`);
    console.log(`With email: ${businesses.filter(b => b.email).length}`);
    console.log(`With phone: ${businesses.filter(b => b.phone || b.mobile).length}`);
    
    await createExcel(businesses, 'kigali_with_google.xlsx');
    console.log('\nDone!');
}

main();