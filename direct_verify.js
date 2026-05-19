const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const fs = require('fs');

async function searchOnDuckDuckGo(businessName) {
    try {
        const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(businessName + ' Rwanda')}`;
        
        const response = await axios.get(url, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            },
            timeout: 15000
        });
        
        const $ = cheerio.load(response.data);
        let results = [];
        
        $('.result').each((i, el) => {
            if (i < 3) {
                const title = $(el).find('.result__title').text().trim();
                const link = $(el).find('.result__url').attr('href') || '';
                const snippet = $(el).find('.result__snippet').text().trim();
                results.push({ title, link, snippet });
            }
        });
        
        return results;
    } catch (e) {
        return [];
    }
}

async function extractWebsiteFromResults(results, businessName) {
    const nameLower = businessName.toLowerCase().replace(/ (ltd|llc|inc|limited|sarl)$/i, '').trim();
    
    for (const r of results) {
        const link = r.link || '';
        const title = r.title || '';
        const snippet = r.snippet || '';
        
        const combined = (title + ' ' + snippet).toLowerCase();
        
        if (link && link.startsWith('http')) {
            const isBusinessLink = combined.includes(nameLower.split(' ')[0].toLowerCase()) ||
                                   combined.includes('official') ||
                                   combined.includes('website');
            
            if (!link.includes('facebook') && !link.includes('twitter') && !link.includes('linkedin') &&
                !link.includes('instagram') && !link.includes('youtube') && !link.includes('google') &&
                !link.includes('amazon') && !link.includes('wikipedia')) {
                
                if (isBusinessLink || r.title.toLowerCase().includes(nameLower.split(' ')[0].toLowerCase())) {
                    return link;
                }
            }
        }
    }
    
    for (const r of results) {
        const link = r.link || '';
        if (link && link.startsWith('http') && 
            !link.includes('facebook') && !link.includes('twitter') && !link.includes('linkedin') &&
            !link.includes('google') && !link.includes('wikipedia')) {
            return link;
        }
    }
    
    return '';
}

async function checkWebsite(url) {
    if (!url) return null;
    try {
        const response = await axios.get(url, {
            timeout: 8000,
            validateStatus: s => s < 500,
            maxRedirects: 3
        });
        return response.status >= 200 && response.status < 400;
    } catch { return false; }
}

async function loadProgress() {
    if (fs.existsSync('progress.json')) {
        return JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    }
    return [];
}

async function main() {
    console.log('Loading data...');
    let businesses = await loadProgress();
    console.log(`Total: ${businesses.length}`);
    
    let verified = 0;
    let working = 0;
    
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        
        if (i % 20 === 0) console.log(`[${i + 1}/${businesses.length}] Checking...`);
        
        if (!b.website && !b.verifiedWebsite) {
            const results = await searchOnDuckDuckGo(b.name);
            
            const website = await extractWebsiteFromResults(results, b.name);
            
            if (website) {
                b.verifiedWebsite = website;
                console.log(`  ${b.name}: Found ${website}`);
                
                b.websiteWorking = await checkWebsite(website);
                console.log(`    Working: ${b.websiteWorking}`);
                
                verified++;
            }
        }
        
        if ((i + 1) % 50 === 0) {
            fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
            console.log(`  Saved. Verified so far: ${verified}`);
        }
        
        await new Promise(r => setTimeout(r, 600));
    }
    
    fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
    
    console.log(`\nTotal verified: ${verified}`);
    
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Verified');
    ws.columns = [
        { header: 'Business Name', key: 'name', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone', key: 'phone', width: 15 },
        { header: 'Website', key: 'website', width: 40 },
        { header: 'Status', key: 'status', width: 12 },
        { header: 'URL', key: 'url', width: 40 }
    ];
    
    for (const b of businesses) {
        ws.addRow({
            name: b.name,
            category: b.category || 'N/A',
            email: b.email || 'N/A',
            phone: b.phone || b.mobile || 'N/A',
            website: b.verifiedWebsite || b.website || 'N/A',
            status: b.websiteWorking === true ? 'Working' : (b.websiteWorking === false ? 'Not Working' : (b.verifiedWebsite ? 'Unknown' : 'N/A')),
            url: b.url || ''
        });
    }
    
    await wb.xlsx.writeFile('kigali_verified.xlsx');
    console.log('Saved: kigali_verified.xlsx');
}

main();