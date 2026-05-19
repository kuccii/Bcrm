const axios = require('axios');
const cheerio = require('cheerio');

async function getCategories() {
    const response = await axios.get('https://www.rwandayp.com/browse-business-directory', {
        headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    const $ = cheerio.load(response.data);
    const categories = [];
    
    $('a[href*="/category/"]').each((i, el) => {
        const name = $(el).text().trim();
        const href = $(el).attr('href');
        if (name && href && name.length > 1 && !name.includes('See All')) {
            categories.push({ name, href });
        }
    });
    
    console.log('Found', categories.length, 'categories');
    categories.forEach(c => console.log(c.name + ' -> ' + c.href));
    
    return categories;
}

getCategories().catch(console.error);