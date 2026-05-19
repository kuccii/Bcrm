const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const fs = require('fs');

const DELAY_MS = 500;
const BATCH_SIZE = 50;
const CHECK_WEBSITE = true;

const categories = {};
const categoriesList = [
    'Business Services', 'Computers & Internet', 'Entertainment & Media', 
    'Events & Conferences', 'Finances & Insurance', 'Food & Drink',
    'Health & Beauty', 'Legal', 'Manufacturing & Industry', 'Shopping',
    'Tourism & Accommodation', 'Tradesmen & Construction', 'Transport & Motoring',
    'Public & Social Services', 'Property'
];

async function fetchPage(url) {
    const response = await axios.get(url, {
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        },
        timeout: 30000
    });
    return cheerio.load(response.data);
}

function getNextPageUrl($, currentUrl, pageNum) {
    const activePage = $('ul.pagination li.active a').text().trim();
    const activeNum = parseInt(activePage) || pageNum;
    
    const nextBtn = $('ul.pagination li.next a').attr('href');
    if (nextBtn) {
        if (nextBtn.startsWith('http')) return nextBtn;
        return 'https://www.rwandayp.com' + nextBtn;
    }
    
    const lastLink = $('ul.pagination a[href*="/page/"]').last().attr('href');
    if (lastLink) {
        const match = lastLink.match(/\/page\/(\d+)/);
        if (match) {
            const maxPage = parseInt(match[1]);
            if (pageNum < maxPage) {
                const base = currentUrl.replace(/\/page\/\d+/, '').replace(/\/\d+$/, '');
                return `${base}/${pageNum + 1}`;
            }
        }
    }
    
    if (pageNum < 150) {
        const base = currentUrl.replace(/\/page\/\d+/, '').replace(/\/\d+$/, '');
        return pageNum === 1 ? `${base}/2` : `${base}/${pageNum + 1}`;
    }
    return null;
}

async function getBusinessLinks(baseUrl, maxPages = 150) {
    const links = [];
    let currentUrl = baseUrl;
    let pageNum = 1;
    
    console.log('Getting business links from all pages...');
    
    while (pageNum <= maxPages) {
        if (pageNum % 20 === 0) console.log(`  Page ${pageNum}...`);
        try {
            const $ = await fetchPage(currentUrl);
            
            $('a[href*="/company/"]').each((i, el) => {
                const name = $(el).text().trim();
                const href = $(el).attr('href');
                const parentText = $(el).closest('div, li, tr').text() || '';
                const categoryMatch = parentText.match(/[A-Za-z\s&]+(?=\s+\d+)/);
                const category = categoryMatch ? categoryMatch[0].trim().split('\n')[0].trim() : '';
                
                if (name && name.length > 2 && name.length < 100 && 
                    href && !name.includes('View') && !name.includes('Profile') &&
                    !name.includes('Enquiry') && !name.includes('Send') &&
                    !name.includes('Get Listed')) {
                    if (!links.find(l => l.name.toLowerCase() === name.toLowerCase())) {
                        links.push({ name, href: 'https://www.rwandayp.com' + href, category: '' });
                    }
                }
            });
            
            const nextUrl = getNextPageUrl($, currentUrl, pageNum);
            if (!nextUrl) break;
            currentUrl = nextUrl;
            pageNum++;
            await new Promise(resolve => setTimeout(resolve, DELAY_MS));
        } catch (e) {
            console.log(`  Error on page ${pageNum}: ${e.message}`);
            break;
        }
    }
    
    console.log(`Found ${links.length} business links`);
    return links;
}

async function getBusinessDetails(url, name) {
    try {
        const $ = await fetchPage(url);
        
        let email = '';
        let phone = '';
        let mobile = '';
        let website = '';
        let address = '';
        let category = '';
        let description = '';
        let contactPerson = '';
        let manager = '';
        
        const pageText = $.html();
        
        $('a[href^="mailto:"]').each((i, el) => {
            const mail = $(el).attr('href').replace('mailto:', '').trim();
            if (mail && !email) email = mail;
        });
        $('a[href^="mailto:"]').each((i, el) => {
            const mail = $(el).text().trim();
            if (mail.includes('@') && !email) email = mail;
        });
        
        $('a[href^="tel:"]').each((i, el) => {
            const tel = $(el).attr('href').replace('tel:', '').trim();
            if (tel && !phone) phone = tel;
        });
        
        $.text().match(/\+250\s?[0-9]{8,9}/g)?.forEach(match => {
            if (!phone && match.length > 8) phone = match;
        });
        $.text().match(/0[0-9]{8,9}/g)?.forEach(match => {
            if (!mobile && match.length === 9) mobile = match;
        });
        
        $('a[href^="http"]:not([href*="rwandayp"])').each((i, el) => {
            const href = $(el).attr('href');
            const text = $(el).text().toLowerCase();
            if (href && !website && !href.includes('google') && !href.includes('facebook') && 
                !href.includes('twitter') && !href.includes('instagram') && 
                !href.includes('linkedin') && !href.includes('youtube') &&
                !href.includes('maps') && !href.includes('goo.gl')) {
                if (text.includes('website') || !href.includes('daddr=') || !href.includes('saddr=')) {
                    website = href.split('?')[0];
                }
            }
        });
        
        const websiteBtn = $('a:contains("Website"), a.btn-primary, a[href*="visit"]:contains("Website")').attr('href');
        if (websiteBtn && websiteBtn.startsWith('http') && !websiteBtn.includes('google')) {
            website = websiteBtn.split('?')[0];
        }
        
        address = $('.address, .location, [class*="address"]').first().text().trim() ||
                  $('span[itemprop="streetAddress"]').text().trim() || '';
        
        category = $('a[href*="/category/"]').first().text().trim() ||
                   $('[class*="category"]').first().text().trim() || '';
        
        description = $('.description, .about, [class*="description"]').first().text().trim() ||
                      $('p').first().text().trim() || '';
        
        contactPerson = $('.contact-person, .contact-name, [class*="contact"]').first().text().trim() || '';
        
        manager = $('.manager, .ceo, .owner, [class*="manager"], [class*="ceo"]').first().text().trim() || '';
        
        if (!manager && contactPerson) manager = contactPerson;
        
        return {
            name,
            url,
            email,
            phone,
            mobile,
            website,
            address,
            category,
            description: description.substring(0, 200),
            contactPerson: manager,
            websiteWorking: null,
            hasWebsite: !!website
        };
        
    } catch (e) {
        return {
            name,
            url,
            email: '', phone: '', mobile: '', website: '', address: '',
            category: '', description: '', contactPerson: '',
            websiteWorking: null, hasWebsite: false,
            error: e.message
        };
    }
}

async function checkWebsiteWorking(url) {
    if (!url || !url.startsWith('http')) return null;
    try {
        const response = await axios.get(url, {
            timeout: 10000,
            validateStatus: (status) => status < 500,
            maxRedirects: 5
        });
        return response.status >= 200 && response.status < 400;
    } catch (e) {
        return false;
    }
}

async function saveProgress(businesses, filename) {
    fs.writeFileSync(filename, JSON.stringify(businesses, null, 2));
}

async function loadProgress(filename) {
    if (fs.existsSync(filename)) {
        return JSON.parse(fs.readFileSync(filename, 'utf8'));
    }
    return [];
}

async function createExcel(businesses, outputPath) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Businesses');
    
    ws.columns = [
        { header: 'Business Name', key: 'name', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Address', key: 'address', width: 35 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone', key: 'phone', width: 15 },
        { header: 'Mobile', key: 'mobile', width: 15 },
        { header: 'Contact Person', key: 'contactPerson', width: 25 },
        { header: 'Website', key: 'website', width: 35 },
        { header: 'Website Working', key: 'websiteWorking', width: 15 },
        { header: 'Description', key: 'description', width: 50 },
        { header: 'Profile URL', key: 'url', width: 40 }
    ];
    
    for (const b of businesses) {
        ws.addRow({
            name: b.name,
            category: b.category || 'N/A',
            address: b.address || 'N/A',
            email: b.email || 'N/A',
            phone: b.phone || 'N/A',
            mobile: b.mobile || 'N/A',
            contactPerson: b.contactPerson || 'N/A',
            website: b.website || 'N/A',
            websiteWorking: b.websiteWorking === null ? 'N/A' : (b.websiteWorking ? 'Yes' : 'No'),
            description: b.description || '',
            url: b.url || ''
        });
    }
    
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Excel saved: ${outputPath}`);
}

async function main() {
    const args = process.argv.slice(2);
    const sourceUrl = args[0] || 'https://www.rwandayp.com/location/Kigali';
    const outputFile = args[1] || 'detailed_results.xlsx';
    const maxBusinesses = parseInt(args[2]) || 100;
    const maxPages = Math.ceil(maxBusinesses / 20);
    const progressFile = 'progress.json';
    
    console.log(`Source: ${sourceUrl}`);
    console.log(`Max businesses: ${maxBusinesses}`);
    console.log('');
    
    let businesses = await loadProgress(progressFile);
    console.log(`Loaded ${businesses.length} from progress`);
    
    const existingNames = new Set(businesses.map(b => b.name));
    
    const businessLinks = await getBusinessLinks(sourceUrl, maxPages);
    
    const newLinks = businessLinks.filter(l => !existingNames.has(l.name)).slice(0, maxBusinesses - businesses.length);
    
    console.log(`Processing ${newLinks.length} new businesses...\n`);
    
    for (let i = 0; i < newLinks.length; i++) {
        const link = newLinks[i];
        console.log(`[${i + 1}/${newLinks.length}] ${link.name}`);
        
        const details = await getBusinessDetails(link.href, link.name);
        
        if (details.hasWebsite && CHECK_WEBSITE) {
            details.websiteWorking = await checkWebsiteWorking(details.website);
            console.log(`  Website: ${details.website} -> ${details.websiteWorking ? 'OK' : 'Failed'}`);
        }
        
        businesses.push(details);
        
        if ((i + 1) % BATCH_SIZE === 0) {
            await saveProgress(businesses, progressFile);
            console.log(`  Progress saved (${businesses.length} businesses)`);
        }
        
        await new Promise(resolve => setTimeout(resolve, DELAY_MS));
    }
    
    await saveProgress(businesses, progressFile);
    await createExcel(businesses, outputFile);
    
    const withWebsite = businesses.filter(b => b.hasWebsite).length;
    const websiteWorks = businesses.filter(b => b.websiteWorking === true).length;
    const withEmail = businesses.filter(b => b.email).length;
    const withPhone = businesses.filter(b => b.phone || b.mobile).length;
    
    console.log('\n=== Summary ===');
    console.log(`Total businesses: ${businesses.length}`);
    console.log(`With website: ${withWebsite}`);
    console.log(`Website working: ${websiteWorks}`);
    console.log(`With email: ${withEmail}`);
    console.log(`With phone: ${withPhone}`);
    console.log('Done!');
}

main();