module.exports = {
    eveBusiness: {
        baseUrl: 'https://www.evebusiness.com',
        searchPath: '/search',
        searchParam: 'q'
    },
    request: {
        timeout: 15000,
        delayBetweenRequests: 1000,
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    },
    selectors: {
        businessList: [
            'table tbody tr',
            '.business-list li',
            '.company-row',
            '.business-item',
            '[data-business]',
            '.listing-item',
            '.company-card'
        ],
        nameSelectors: [
            'td:first',
            '.name',
            '.company-name',
            '.business-name',
            '[class*="name"]',
            'a[href*="business"]',
            'a[href*="company"]'
        ]
    }
};