const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const fs = require('fs');

const SEARCH_ENGINES = [
    { name: 'DuckDuckGo', baseUrl: 'https://duckduckgo.com/', searchParam: 'q' },
    { name: 'Bing', baseUrl: 'https://www.bing.com/search', searchParam: 'q' }
];

async function searchBusiness(engine, businessName) {
    try {
        const searchUrl = `${engine.baseUrl}${engine.searchParam}=${encodeURIComponent(businessName + ' Rwanda')}`;
        
        const response = await axios.get(searchUrl, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.9',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache'
            },
            timeout: 15000,
            maxRedirects: 5
        });
        
        const $ = cheerio.load(response.data);
        let website = '';
        let description = '';
        
        if (engine.name === 'DuckDuckGo') {
            $('a[data-testid="result-extras-domain-link"]').each((i, el) => {
                if (!website && i === 0) {
                    const href = $(el).attr('href');
                    if (href && href.startsWith('http') && !href.includes('duckduckgo') && !href.includes('yahoo.com')) {
                        website = href.split('?')[0];
                    }
                }
            });
            
            if (!website) {
                $('a.result__a').each((i, el) => {
                    const href = $(el).attr('href');
                    const text = $(el).text().toLowerCase();
                    if (href && !website && (text.includes(businessName.toLowerCase().split(' ')[0]) || text.includes('website'))) {
                        const match = href.match(/uddg=([^&]+)/);
                        if (match) {
                            website = decodeURIComponent(match[1]).split('?')[0];
                        }
                    }
                });
            }
            
            $('a.result__a').first().each((i, el) => {
                const href = $(el).attr('href');
                if (href && href.includes('uddg=')) {
                    const match = href.match(/uddg=([^&]+)/);
                    if (match && !website) {
                        const decoded = decodeURIComponent(match[1]);
                        if (decoded.startsWith('http') && !decoded.includes('duckduckgo')) {
                            website = decoded.split('?')[0];
                        }
                    }
                }
            });
            
            description = $('a.result__a').first().text();
        } 
        else if (engine.name === 'Bing') {
            $('a[href^="http"]:not([href*="bing"]):not([href*="microsoft"])').each((i, el) => {
                const href = $(el).attr('href');
                const text = $(el).text().toLowerCase();
                if (!website && href && href.startsWith('http')) {
                    if (text.includes('website') || text.includes('.com') || text.includes('.rw') || text.includes('.org')) {
                        if (!href.includes('facebook') && !href.includes('twitter') && !href.includes('linkedin')) {
                            website = href.split('?')[0];
                        }
                    }
                }
            });
            
            $('cite').first().each((i, el) => {
                const text = $(el).text();
                if (!website && text.startsWith('http')) {
                    website = text.split('?')[0];
                }
            });
            
            description = $('h2 a').first().text();
        }
        
        return { website, description };
    } catch (e) {
        return { website: '', description: '', error: e.message };
    }
}

async function checkWebsite(url) {
    if (!url || !url.startsWith('http')) return null;
    try {
        const response = await axios.get(url, {
            timeout: 10000,
            validateStatus: (status) => status < 500,
            maxRedirects: 5,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
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
    const worksheet = workbook.addWorksheet('Verified Businesses');
    
    worksheet.columns = [
        { header: 'Business Name', key: 'name', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Address', key: 'address', width: 35 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone', key: 'phone', width: 15 },
        { header: 'Mobile', key: 'mobile', width: 15 },
        { header: 'Contact Person', key: 'contactPerson', width: 25 },
        { header: 'Website', key: 'website', width: 40 },
        { header: 'Website Verified', key: 'websiteVerified', width: 15 },
        { header: 'Status', key: 'status', width: 12 },
        { header: 'Source', key: 'source', width: 15 },
        { header: 'Profile URL', key: 'url', width: 40 }
    ];
    
    for (const b of businesses) {
        const hasWebsite = b.verifiedWebsite || b.website;
        const isWorking = b.websiteWorking;
        const source = b.verifiedWebsite ? 'Verified' : (b.website ? 'RwandaYP' : 'None');
        
        worksheet.addRow({
            name: b.name,
            category: b.category || 'N/A',
            address: b.address || 'N/A',
            email: b.email || 'N/A',
            phone: b.phone || 'N/A',
            mobile: b.mobile || 'N/A',
            contactPerson: b.contactPerson || 'N/A',
            website: hasWebsite || 'N/A',
            websiteVerified: hasWebsite ? 'Yes' : 'No',
            status: isWorking === true ? 'Working' : (isWorking === false ? 'Not Working' : 'N/A'),
            source: source,
            url: b.url || ''
        });
    }
    
    await workbook.xlsx.writeFile(outputPath);
    console.log(`\nExcel saved: ${outputPath}`);
}

async function main() {
    console.log('Loading business data...');
    let businesses = await loadProgress();
    console.log(`Total businesses: ${businesses.length}`);
    
    console.log('\n=== Starting Website Verification ===\n');
    
    let verified = 0;
    let totalWithWebsite = 0;
    let working = 0;
    
    for (let i = 0; i < businesses.length; i++) {
        const b = businesses[i];
        
        console.log(`[${i + 1}/${businesses.length}] ${b.name}`);
        
        if (!b.verifiedWebsite && !b.website) {
            for (const engine of SEARCH_ENGINES) {
                if (b.verifiedWebsite) break;
                
                const result = await searchBusiness(engine, b.name);
                
                if (result.website) {
                    b.verifiedWebsite = result.website;
                    console.log(`  Found via ${engine.name}: ${result.website}`);
                    
                    b.websiteWorking = await checkWebsite(result.website);
                    console.log(`  Status: ${b.websiteWorking === true ? 'WORKING' : (b.websiteWorking === false ? 'NOT WORKING' : 'UNKNOWN')}`);
                    
                    break;
                }
                
                await new Promise(r => setTimeout(r, 500));
            }
        }
        
        if (b.verifiedWebsite || b.website) {
            totalWithWebsite++;
            if (b.websiteWorking === true) working++;
        }
        
        if ((i + 1) % 50 === 0) {
            await saveProgress(businesses);
            console.log(`\n  Progress saved - ${totalWithWebsite} with website, ${working} working\n`);
        }
        
        await new Promise(r => setTimeout(r, 800));
    }
    
    await saveProgress(businesses);
    
    console.log('\n=== FINAL RESULTS ===');
    console.log(`Total businesses: ${businesses.length}`);
    console.log(`With website (verified + RwandaYP): ${totalWithWebsite}`);
    console.log(`Website working: ${working}`);
    console.log(`With email: ${businesses.filter(b => b.email).length}`);
    console.log(`With phone: ${businesses.filter(b => b.phone || b.mobile).length}`);
    console.log(`With category: ${businesses.filter(b => b.category).length}`);
    
    await createExcel(businesses, 'kigali_verified.xlsx');
    console.log('\nDone!');
}

main();