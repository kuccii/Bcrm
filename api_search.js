const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const fs = require('fs');

function generatePossibleDomains(businessName) {
    const domains = [];
    const cleanName = businessName
        .replace(/\s+(Ltd|LLC|Inc|Limited|SARL|PLC|Co\.|Company|Pvt|Group|Holdings|Solutions|Services|Consulting|Tours|Travel|Agency|Hotel|Restaurant|Pharmacy|Clinic|Hospital|School|College|University|Institute|Foundation|Association|Center|Centre)\.?$/gi, '')
        .replace(/[^a-zA-Z0-9\s]/g, '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
    
    const words = businessName.split(' ').filter(w => w.length > 2);
    const mainWord = words[0] ? words[0].toLowerCase().replace(/[^a-z]/g, '') : cleanName;
    
    // Common TLDs for Rwanda
    const tlds = ['.com', '.rw', '.org', '.net'];
    
    // Generate variations
    const variations = [
        cleanName,
        mainWord,
        cleanName.replace(/rwanda$/, ''),
        mainWord + 'rwanda',
        cleanName + 'rwanda'
    ];
    
    // Remove duplicates and empty
    const unique = [...new Set(variations.filter(v => v.length > 2))];
    
    for (const name of unique) {
        for (const tld of tlds) {
            domains.push(`https://www.${name}${tld}`);
            domains.push(`https://${name}${tld}`);
        }
    }
    
    return domains;
}

async function checkDomain(url) {
    try {
        const response = await axios.get(url, {
            timeout: 8000,
            validateStatus: (status) => status < 500,
            maxRedirects: 3,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
            }
        });
        
        if (response.status >= 200 && response.status < 400) {
            const $ = cheerio.load(response.data);
            const title = $('title').text().toLowerCase();
            return { working: true, title: title.substring(0, 60) };
        }
        return { working: false };
    } catch (e) {
        return { working: false, error: e.message };
    }
}

async function searchBusinessWebsite(businessName) {
    const domains = generatePossibleDomains(businessName);
    
    for (const domain of domains.slice(0, 8)) {
        const result = await checkDomain(domain);
        if (result.working) {
            return { found: true, url: domain, title: result.title };
        }
    }
    
    return { found: false };
}

async function loadProgress() {
    if (fs.existsSync('progress.json')) {
        return JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    }
    return [];
}

async function main() {
    console.log('Loading business data...');
    let businesses = await loadProgress();
    console.log(`Total: ${businesses.length}`);
    
    console.log('\n=== Searching for business websites ===\n');
    
    let found = 0;
    let checked = 0;
    
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        
        if (!b.website && !b.verifiedWebsite && !b.domainChecked) {
            checked++;
            console.log(`[${i + 1}/${businesses.length}] ${b.name}`);
            
            const result = await searchBusinessWebsite(b.name);
            
            if (result.found) {
                b.verifiedWebsite = result.url;
                b.websiteTitle = result.title;
                b.websiteWorking = true;
                console.log(`  FOUND: ${result.url}`);
                console.log(`  Title: ${result.title}`);
                found++;
            }
            
            b.domainChecked = true;
            
            if ((i + 1) % 50 === 0) {
                fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
                console.log(`\n  Progress saved. Found: ${found}/${checked}\n`);
            }
            
            await new Promise(r => setTimeout(r, 300));
        }
    }
    
    fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
    
    console.log(`\n=== RESULTS ===`);
    console.log(`Total checked: ${checked}`);
    console.log(`Websites found: ${found}`);
    
    // Create Excel
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Kigali Businesses');
    
    ws.columns = [
        { header: 'No', key: 'no', width: 5 },
        { header: 'Business Name', key: 'name', width: 32 },
        { header: 'Category', key: 'category', width: 28 },
        { header: 'Email', key: 'email', width: 32 },
        { header: 'Phone', key: 'phone', width: 18 },
        { header: 'Website', key: 'website', width: 40 },
        { header: 'Status', key: 'status', width: 12 },
        { header: 'URL', key: 'url', width: 45 }
    ];
    
    let no = 1;
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || 'N/A';
        const status = b.websiteWorking === true ? 'Working' : (website !== 'N/A' ? 'Not Working' : 'N/A');
        
        ws.addRow({
            no: no++,
            name: b.name,
            category: b.category || 'N/A',
            email: b.email || 'N/A',
            phone: b.phone || b.mobile || 'N/A',
            website: website,
            status: status,
            url: b.url || ''
        });
    }
    
    await wb.xlsx.writeFile('Kigali_Business_Directory.xlsx');
    console.log('\nSaved: Kigali_Business_Directory.xlsx');
    console.log('Done!');
}

main();