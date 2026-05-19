const axios = require('axios');
const cheerio = require('cheerio');

async function test() {
    try {
        const url = 'https://html.duckduckgo.com/html/?q=NanoTranslate Rwanda';
        const response = await axios.get(url, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            },
            timeout: 15000
        });
        
        const $ = cheerio.load(response.data);
        console.log('Page title:', $('title').text());
        console.log('Results found:', $('.result').length);
        
        $('.result').each((i, el) => {
            if (i < 5) {
                const title = $(el).find('.result__title').text().trim();
                const url = $(el).find('.result__url').attr('href') || '';
                console.log(`\nResult ${i + 1}:`);
                console.log('Title:', title.substring(0, 60));
                console.log('URL:', url.substring(0, 80));
            }
        });
        
        console.log('\n\n--- Raw body sample ---');
        console.log(response.data.substring(0, 2000));
        
    } catch (e) {
        console.log('Error:', e.message);
    }
}

test();