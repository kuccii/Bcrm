const axios = require('axios');
const cheerio = require('cheerio');

axios.get('https://www.rwandayp.com/location/Kigali', {
    headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
}).then(res => {
    const $ = cheerio.load(res.data);
    
    console.log('=== Debug Output ===');
    console.log('H2 count:', $('h2').length);
    console.log('H3 count:', $('h3').length);
    console.log('H4 count:', $('h4').length);
    console.log('A with /company/:', $('a[href*="/company/"]').length);
    console.log('Row company:', $('.row.company').length);
    console.log('Business item:', $('.business-item').length);
    console.log('Col-md-4:', $('.col-md-4').length);
    console.log('First h2:', $('h2').first().text().trim());
    console.log('First company link:', $('a[href*="/company/"]').first().text().trim());
    console.log('---');
    console.log($('.row.company').first().html());
}).catch(e => console.log(e.message));