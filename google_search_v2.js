const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs');

async function searchBusiness(name) {
    try {
        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(name + " Rwanda site:maps.google.com")}`;
        const response = await axios.get(searchUrl, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5'
            },
            timeout: 10000,
            maxRedirects: 5
        });
        
        const html = response.data;
        let website = '';
        let mapsUrl = '';
        
        const mapMatch = html.match(/https:\/\/maps\.google\.com\/[^"'\s]+/);
        if (mapMatch) mapsUrl = mapMatch[0].split('&')[0];
        
        const websiteMatch = html.match(/https:\/\/[^"'\s<>]+\.(com|rw|org|net)[^"'\s<>]*/);
        if (websiteMatch && !websiteMatch[0].includes('google') && !websiteMatch[0].includes('facebook')) {
            website = websiteMatch[0].split('&')[0];
        }
        
        return { website, mapsUrl };
    } catch (e) {
        return { website: '', mapsUrl: '', error: e.message };
    }
}

async function checkWebsite(url) {
    if (!url) return false;
    try {
        const response = await axios.get(url, { timeout: 8000, validateStatus: s => s < 500 });
        return response.status >= 200 && response.status < 400;
    } catch { return false; }
}

async function main() {
    console.log('Loading data...');
    let businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    console.log(`Total: ${businesses.length}`);
    
    const needSearch = businesses.filter(b => !b.website && !b.googleSearched);
    console.log(`Need Google search: ${needSearch.length}`);
    
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        
        if (!b.website && !b.googleSearched) {
            console.log(`[${i+1}/${businesses.length}] ${b.name}`);
            
            const result = await searchBusiness(b.name);
            
            if (result.website) {
                b.googleWebsite = result.website;
                b.googleWebsiteWorking = await checkWebsite(result.website);
                console.log(`  Website: ${result.website} (${b.googleWebsiteWorking ? 'OK' : 'Failed'})`);
            }
            if (result.mapsUrl) b.googleMapsUrl = result.mapsUrl;
            
            b.googleSearched = true;
            
            if ((i + 1) % 20 === 0) {
                fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
                console.log('  Saved');
            }
            
            await new Promise(r => setTimeout(r, 800));
        }
    }
    
    fs.writeFileSync('progress.json', JSON.stringify(businesses, null, 2));
    
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Results');
    ws.columns = [
        { header: 'Business Name', key: 'name', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone', key: 'phone', width: 15 },
        { header: 'Website', key: 'website', width: 35 },
        { header: 'Website Source', key: 'source', width: 12 },
        { header: 'Working', key: 'working', width: 10 },
        { header: 'Profile URL', key: 'url', width: 40 }
    ];
    
    for (const b of businesses) {
        ws.addRow({
            name: b.name,
            category: b.category || 'N/A',
            email: b.email || 'N/A',
            phone: b.phone || b.mobile || 'N/A',
            website: b.googleWebsite || b.website || 'N/A',
            source: b.googleWebsite ? 'Google' : (b.website ? 'RwandaYP' : 'N/A'),
            working: b.googleWebsiteWorking === true ? 'Yes' : (b.googleWebsiteWorking === false ? 'No' : (b.websiteWorking === true ? 'Yes' : (b.websiteWorking === false ? 'No' : 'N/A'))),
            url: b.url || ''
        });
    }
    
    await wb.xlsx.writeFile('kigali_final.xlsx');
    console.log('\nDone! Saved: kigali_final.xlsx');
    
    const hasWeb = businesses.filter(b => b.website || b.googleWebsite).length;
    const works = businesses.filter(b => b.websiteWorking === true || b.googleWebsiteWorking === true).length;
    console.log(`\nSummary: ${hasWeb} with website, ${works} working`);
}

main();