const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');

const MAX_PAGES = 150;
const DELAY_MS = 200;

async function fetchPage(url) {
    const response = await axios.get(url, {
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        },
        timeout: 30000
    });
    return cheerio.load(response.data);
}

function getNextPageUrl($, currentUrl, pageNum) {
    const paginationLinks = $('ul.pagination a[href*="/page/"]');
    if (paginationLinks.length > 0) {
        const lastLink = paginationLinks.last().attr('href');
        const match = lastLink.match(/\/page\/(\d+)/);
        if (match) {
            const maxPage = parseInt(match[1]);
            if (pageNum < maxPage) {
                return currentUrl.replace(/\/\d+$/, '') + `/${pageNum + 1}`;
            }
        }
    }
    
    const nextBtn = $('ul.pagination li.next a, a.next').attr('href');
    if (nextBtn) {
        if (nextBtn.startsWith('http')) return nextBtn;
        return new URL(nextBtn, currentUrl).href;
    }
    
    if (pageNum < 150) {
        return currentUrl.replace(/\/\d+$/, '') + `/${pageNum + 1}`;
    }
    
    return null;
}

async function fetchAllBusinesses(baseUrl) {
    const allBusinesses = [];
    const seen = new Set();
    let currentUrl = baseUrl;
    let pageNum = 1;
    
    while (pageNum <= MAX_PAGES) {
        console.log(`Page ${pageNum}`);
        
        try {
            const $ = await fetchPage(currentUrl);
            
            const companyLinks = $('a[href*="/company/"]');
            const businessMap = new Map();
            
            companyLinks.each((i, el) => {
                const name = $(el).text().trim();
                const href = $(el).attr('href') || '';
                
                if (name && name.length > 2 && name.length < 100 && 
                    href.includes('/company/') && !name.includes('View') && 
                    !name.includes('Profile') && !name.includes('Enquiry') &&
                    !name.includes('Send') && !name.includes('Get Listed') &&
                    !name.includes('Website') && !name.includes('E-mail') &&
                    !name.includes('Map')) {
                    
                    if (!businessMap.has(name)) {
                        const parentText = $(el).closest('div, li, tr').text().toLowerCase();
                        const hasWebsite = parentText.includes('website');
                        
                        businessMap.set(name, {
                            name: name,
                            hasWebsite: hasWebsite
                        });
                    }
                }
            });
            
            let found = 0;
            businessMap.forEach((value, key) => {
                if (!seen.has(key.toLowerCase())) {
                    seen.add(key.toLowerCase());
                    allBusinesses.push({
                        name: value.name,
                        website: '',
                        hasWebsite: value.hasWebsite,
                        working: false
                    });
                    found++;
                }
            });
            
            console.log(`  Found ${found} (total: ${allBusinesses.length})`);
            
            const nextUrl = getNextPageUrl($, currentUrl, pageNum);
            if (!nextUrl) break;
            
            currentUrl = nextUrl;
            pageNum++;
            await new Promise(resolve => setTimeout(resolve, DELAY_MS));
            
        } catch (error) {
            console.log(`  Error: ${error.message}`);
            break;
        }
    }
    
    return allBusinesses;
}

async function createExcel(businesses, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Results');
    
    worksheet.columns = [
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Has Website', key: 'hasWebsite', width: 12 }
    ];
    
    const hasWebsite = businesses.filter(b => b.hasWebsite).length;
    console.log(`\nBusinesses with website: ${hasWebsite}`);
    
    for (const b of businesses) {
        worksheet.addRow({
            name: b.name,
            hasWebsite: b.hasWebsite ? 'Yes' : 'No'
        });
    }
    
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Saved: ${outputPath}`);
    console.log(`Total: ${businesses.length} businesses`);
}

function parseArgs() {
    const args = process.argv.slice(2);
    if (args.length === 0) {
        console.log('Usage: node index.js <url> [output]');
        process.exit(1);
    }
    return { sourceUrl: args[0], outputFile: args[1] || 'results.xlsx' };
}

async function main() {
    const { sourceUrl, outputFile } = parseArgs();
    
    console.log('Scanning businesses...\n');
    const businesses = await fetchAllBusinesses(sourceUrl);
    
    console.log(`\nFound ${businesses.length} businesses`);
    await createExcel(businesses, outputFile);
    console.log('Done!');
}

main();