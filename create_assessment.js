const ExcelJS = require('exceljs');
const fs = require('fs');

(async function() {
    const businesses = JSON.parse(fs.readFileSync('progress.json', 'utf8'));
    
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Business Assessment System';
    wb.created = new Date();
    
    const hdr = {
        font: { bold: true, color: { argb: 'FFFFFF' }, size: 10 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1565C0' } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: { bottom: { style: 'medium', color: { argb: '0D47A1' } } }
    };

    // ======== SERVICE DEFINITIONS ========
    // Service categories with industry-specific needs
    const serviceCatalog = {
        'Website': { base: 350, desc: 'Professional website with mobile optimization, contact forms, Google Maps integration' },
        'E-commerce': { base: 800, desc: 'Online store with product catalog, cart, payment gateway (Paystack, etc.)' },
        'Booking System': { base: 600, desc: 'Online booking/reservation calendar, automated confirmations, SMS reminders' },
        'POS System': { base: 500, desc: 'Point of sale, inventory tracking, sales reports, receipt printing' },
        'CRM': { base: 400, desc: 'Customer management, follow-up automation, deal tracking, email integration' },
        'Inventory Mgmt': { base: 450, desc: 'Stock tracking, supplier management, low stock alerts, barcode scanning' },
        'Accounting/ERP': { base: 700, desc: 'Invoice generation, expense tracking, tax reports, bank reconciliation' },
        'Marketing Automation': { base: 350, desc: 'Email campaigns, SMS blasts, social media scheduling, analytics' },
        'Communication': { base: 300, desc: 'VOIP phone system, WhatsApp Business API, chatbot, auto-responders' },
        'Online Ordering': { base: 500, desc: 'Menu display, order placement, payment, delivery tracking' },
        'HR & Payroll': { base: 400, desc: 'Employee management, attendance, payroll calculation, leave tracking' },
        'Analytics & Reporting': { base: 300, desc: 'Dashboard, KPI tracking, custom reports, data visualization' },
        'Security Systems': { base: 500, desc: 'CCTV, access control, alarm systems, remote monitoring' },
        'Cloud Backup': { base: 200, desc: 'Automated cloud backup, data recovery, file sync, secure storage' },
        'Social Media Mgmt': { base: 250, desc: 'Content creation, posting schedule, engagement tracking, ad management' },
        'SMS/Notification': { base: 200, desc: 'Bulk SMS, appointment reminders, promotional campaigns, alerts' },
        'Payment Gateway': { base: 300, desc: 'Mobile money (M-Pesa, Airtel), card payments, invoice links' },
        'Delivery Logistics': { base: 400, desc: 'Route optimization, dispatch tracking, delivery confirmation, fleet mgmt' },
        'Learning Mgmt': { base: 500, desc: 'Online courses, student portal, assignment tracking, certificates' },
        'Patient Mgmt': { base: 600, desc: 'Patient records, appointment scheduling, prescriptions, lab integration' },
    };

    // ======== INDUSTRY SERVICE MAP ========
    // Maps each category to relevant services they need
    const industryNeeds = {
        'Restaurants': ['Website', 'Online Ordering', 'Booking System', 'POS System', 'Inventory Mgmt', 'Marketing Automation', 'Social Media Mgmt'],
        'Hotels': ['Website', 'Booking System', 'POS System', 'CRM', 'Marketing Automation', 'Communication', 'Analytics & Reporting'],
        'Tour Operators': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Payment Gateway', 'Communication', 'Social Media Mgmt'],
        'Travel Agents': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Communication', 'Payment Gateway'],
        'Doctors and Clinics': ['Website', 'Patient Mgmt', 'Booking System', 'CRM', 'SMS/Notification', 'Cloud Backup'],
        'Pharmacies': ['Website', 'Inventory Mgmt', 'POS System', 'SMS/Notification', 'Accounting/ERP', 'Online Ordering'],
        'Fashion': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'POS System'],
        'Food Retailers': ['Website', 'Online Ordering', 'POS System', 'Inventory Mgmt', 'Delivery Logistics', 'Marketing Automation'],
        'Retail Services': ['Website', 'E-commerce', 'POS System', 'Inventory Mgmt', 'Accounting/ERP', 'Marketing Automation'],
        'Education': ['Website', 'Learning Mgmt', 'SMS/Notification', 'CRM', 'Cloud Backup', 'Communication'],
        'RealEstate Agents': ['Website', 'CRM', 'Marketing Automation', 'Social Media Mgmt', 'Communication', 'Analytics & Reporting'],
        'Consultants': ['Website', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup', 'Accounting/ERP'],
        'Consulting': ['Website', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup', 'Accounting/ERP'],
        'Lawyers': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP', 'Booking System'],
        'Automotive': ['Website', 'Inventory Mgmt', 'POS System', 'CRM', 'SMS/Notification', 'Accounting/ERP'],
        'Construction': ['Website', 'CRM', 'Accounting/ERP', 'Cloud Backup', 'Analytics & Reporting', 'Communication'],
        'Beauty Professionals': ['Website', 'Booking System', 'Social Media Mgmt', 'Marketing Automation', 'POS System', 'SMS/Notification'],
        'Catering': ['Website', 'Online Ordering', 'Delivery Logistics', 'Inventory Mgmt', 'Marketing Automation', 'Communication'],
        'Cafes': ['Website', 'Online Ordering', 'POS System', 'Inventory Mgmt', 'Marketing Automation', 'Social Media Mgmt'],
        'Pharmacies': ['Website', 'Inventory Mgmt', 'POS System', 'SMS/Notification', 'Online Ordering', 'Cloud Backup'],
        'Hardware Stores': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Accounting/ERP', 'Delivery Logistics'],
        'Energy Suppliers': ['Website', 'CRM', 'Accounting/ERP', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Industrial Services': ['Website', 'CRM', 'Inventory Mgmt', 'Accounting/ERP', 'Cloud Backup', 'Communication'],
        'General Business': ['Website', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup', 'Accounting/ERP'],
        'Advertising': ['Website', 'Social Media Mgmt', 'Marketing Automation', 'Analytics & Reporting', 'CRM', 'Communication'],
        'Associations': ['Website', 'CRM', 'Communication', 'Marketing Automation', 'Cloud Backup', 'SMS/Notification'],
        'General Office Services': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP', 'Analytics & Reporting'],
        'Event Services': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Social Media Mgmt', 'Communication'],
        'Dry Cleaning': ['Website', 'Online Ordering', 'POS System', 'Inventory Mgmt', 'Delivery Logistics', 'SMS/Notification'],
        'Communications': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Cloud Backup', 'Communication'],
        'Medical Equipment': ['Website', 'Inventory Mgmt', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP'],
        'Market Research': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Cloud Backup', 'Communication'],
        'Secretarial Services': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP', 'Marketing Automation'],
        'Small Business': ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication', 'Accounting/ERP'],
        'Business Centres': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP', 'Security Systems'],
        'Banks, Credit Unions': ['Website', 'CRM', 'Accounting/ERP', 'Security Systems', 'Analytics & Reporting', 'Cloud Backup'],
        'Air Travel': ['Website', 'Booking System', 'CRM', 'Communication', 'Marketing Automation', 'Analytics & Reporting'],
        'Leisure': ['Website', 'Booking System', 'Social Media Mgmt', 'Marketing Automation', 'POS System', 'Communication'],
        'Printing': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Accounting/ERP', 'Delivery Logistics'],
        'Software Applications': ['Website', 'CRM', 'Cloud Backup', 'Marketing Automation', 'Analytics & Reporting', 'Communication'],
        'Logistics': ['Website', 'Delivery Logistics', 'CRM', 'Analytics & Reporting', 'Communication', 'Accounting/ERP'],
        'Architectural Services': ['Website', 'Cloud Backup', 'CRM', 'Communication', 'Analytics & Reporting', 'Accounting/ERP'],
        'Web Design': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Cloud Backup', 'Communication'],
        'Web Development': ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Analytics & Reporting', 'Communication'],
        'Arts and Crafts': ['Website', 'E-commerce', 'Social Media Mgmt', 'Marketing Automation', 'Inventory Mgmt', 'Delivery Logistics'],
        'Car Parts and Accessories': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Delivery Logistics', 'Marketing Automation'],
        'Building Materials': ['Website', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Accounting/ERP', 'Communication'],
        'Civil Engineering': ['Website', 'CRM', 'Accounting/ERP', 'Cloud Backup', 'Analytics & Reporting', 'Communication'],
        'Financial Activity': ['Website', 'CRM', 'Accounting/ERP', 'Security Systems', 'Analytics & Reporting', 'Cloud Backup'],
        'Attractions': ['Website', 'Booking System', 'POS System', 'Marketing Automation', 'Social Media Mgmt', 'Analytics & Reporting'],
        'Cleaning Equipment & Services': ['Website', 'CRM', 'Inventory Mgmt', 'Marketing Automation', 'Delivery Logistics', 'Communication'],
        'Information Technology': ['Website', 'Cloud Backup', 'Security Systems', 'CRM', 'Analytics & Reporting', 'Communication'],
        'Security Services': ['Website', 'Security Systems', 'CRM', 'Communication', 'Cloud Backup', 'SMS/Notification'],
        'Computer Software Solution': ['Website', 'Cloud Backup', 'CRM', 'Analytics & Reporting', 'Marketing Automation', 'Communication'],
        'Beauty Products': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'POS System'],
        'Car Rental': ['Website', 'Booking System', 'CRM', 'Communication', 'Marketing Automation', 'Analytics & Reporting'],
        'Industrial Equipment': ['Website', 'Inventory Mgmt', 'CRM', 'Accounting/ERP', 'Delivery Logistics', 'Cloud Backup'],
        'Training': ['Website', 'Learning Mgmt', 'CRM', 'Marketing Automation', 'SMS/Notification', 'Cloud Backup'],
        'Home and Garden': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Marketing Automation'],
        'Investment Companies': ['Website', 'CRM', 'Analytics & Reporting', 'Security Systems', 'Cloud Backup', 'Communication'],
        'Legal Services': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Accounting/ERP', 'Booking System'],
        'Pubs and Clubs': ['Website', 'POS System', 'Booking System', 'Inventory Mgmt', 'Marketing Automation', 'Social Media Mgmt'],
        'Film, Television and Video': ['Website', 'Cloud Backup', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication'],
        'Specialist Accommodation': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Communication', 'POS System'],
        'Sports': ['Website', 'Booking System', 'CRM', 'POS System', 'Marketing Automation', 'Social Media Mgmt'],
        'Clothing and Accessories': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'POS System'],
        'Brokers': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Tourist Information': ['Website', 'CRM', 'Communication', 'Marketing Automation', 'Analytics & Reporting', 'Social Media Mgmt'],
        'Banking Equipment': ['Website', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Accounting/ERP', 'Communication'],
        'Food Manufacturing': ['Website', 'Inventory Mgmt', 'Accounting/ERP', 'Delivery Logistics', 'CRM', 'Cloud Backup'],
        'Health and Safety': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Marketing Automation', 'Analytics & Reporting'],
        'Online Content': ['Website', 'Marketing Automation', 'Social Media Mgmt', 'Analytics & Reporting', 'CRM', 'Cloud Backup'],
        'Pets and Animals': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'CRM', 'Marketing Automation'],
        'Audit and Accounting': ['Website', 'Accounting/ERP', 'CRM', 'Cloud Backup', 'Security Systems', 'Analytics & Reporting'],
        'Colleges': ['Website', 'Learning Mgmt', 'SMS/Notification', 'CRM', 'Cloud Backup', 'Communication'],
        'Fire Safety Consultants': ['Website', 'CRM', 'Communication', 'Cloud Backup', 'Marketing Automation', 'Analytics & Reporting'],
        'Shops': ['Website', 'E-commerce', 'POS System', 'Inventory Mgmt', 'Marketing Automation', 'Social Media Mgmt'],
        'Electronic Equipment': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Cloud Backup'],
        'Bed and Breakfast': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Communication', 'POS System'],
        'Event Equipment': ['Website', 'Booking System', 'Inventory Mgmt', 'CRM', 'Marketing Automation', 'Communication'],
        'Translation Services': ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication', 'Analytics & Reporting'],
        'Books': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Marketing Automation', 'Social Media Mgmt'],
        'Music': ['Website', 'E-commerce', 'Social Media Mgmt', 'Marketing Automation', 'CRM', 'Cloud Backup'],
        'Public Relations': ['Website', 'Social Media Mgmt', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication'],
        'FinTech': ['Website', 'CRM', 'Security Systems', 'Analytics & Reporting', 'Cloud Backup', 'Marketing Automation'],
        'Conferences': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Communication', 'Analytics & Reporting'],
        'Insurance Companies': ['Website', 'CRM', 'Analytics & Reporting', 'Security Systems', 'Cloud Backup', 'Communication'],
        'Camping and Caravans': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Social Media Mgmt', 'Communication'],
        'Shipping & Port Agent': ['Website', 'CRM', 'Delivery Logistics', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Web Services': ['Website', 'Cloud Backup', 'Security Systems', 'CRM', 'Analytics & Reporting', 'Marketing Automation'],
        'Business Development': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'AgTech': ['Website', 'Inventory Mgmt', 'CRM', 'Analytics & Reporting', 'Cloud Backup', 'Communication'],
        'Internet Service Providers': ['Website', 'CRM', 'Analytics & Reporting', 'Communication', 'Cloud Backup', 'Marketing Automation'],
        'Air Transport': ['Website', 'Booking System', 'CRM', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Industrial Supplies': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Accounting/ERP'],
        'Computer Services': ['Website', 'Cloud Backup', 'CRM', 'Security Systems', 'Analytics & Reporting', 'Communication'],
        'Hairdressers': ['Website', 'Booking System', 'Social Media Mgmt', 'POS System', 'Marketing Automation', 'SMS/Notification'],
        'Ministries': ['Website', 'CRM', 'Cloud Backup', 'Analytics & Reporting', 'Communication', 'Security Systems'],
        'Transport': ['Website', 'Delivery Logistics', 'CRM', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Interior Design': ['Website', 'Social Media Mgmt', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication'],
        'Computer Consumables': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Delivery Logistics', 'CRM'],
        'Catering Equipment': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Accounting/ERP'],
        'Security': ['Website', 'Security Systems', 'CRM', 'Cloud Backup', 'Communication', 'SMS/Notification'],
        'Eco-friendly Products': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'Delivery Logistics'],
        'Sales Outsourcing': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Photography': ['Website', 'Social Media Mgmt', 'CRM', 'Marketing Automation', 'Cloud Backup', 'E-commerce'],
        'Taxis': ['Website', 'Booking System', 'Communication', 'Delivery Logistics', 'SMS/Notification', 'Marketing Automation'],
        'Web Hosting': ['Website', 'Cloud Backup', 'CRM', 'Security Systems', 'Analytics & Reporting', 'Marketing Automation'],
        'Apartments': ['Website', 'Booking System', 'CRM', 'Communication', 'Marketing Automation', 'Cloud Backup'],
        'Water Treatment': ['Website', 'CRM', 'Inventory Mgmt', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Cleaning': ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'SMS/Notification', 'Accounting/ERP'],
        'Weddings': ['Website', 'Social Media Mgmt', 'CRM', 'Marketing Automation', 'E-commerce', 'Booking System'],
        'Networking': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Bicycles': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'CRM', 'Marketing Automation'],
        'Jewellery': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'POS System'],
        'Project Management': ['Website', 'CRM', 'Analytics & Reporting', 'Cloud Backup', 'Marketing Automation', 'Communication'],
        'Electrical Service': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Marketing Automation', 'Accounting/ERP'],
        'Commercial Property': ['Website', 'CRM', 'Analytics & Reporting', 'Marketing Automation', 'Communication', 'Cloud Backup'],
        'Schools': ['Website', 'Learning Mgmt', 'SMS/Notification', 'CRM', 'Cloud Backup', 'Communication'],
        'Tax Consultants': ['Website', 'Accounting/ERP', 'CRM', 'Cloud Backup', 'Marketing Automation', 'Communication'],
        'Engineering': ['Website', 'CRM', 'Analytics & Reporting', 'Cloud Backup', 'Communication', 'Accounting/ERP'],
        'Electrical Goods': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Delivery Logistics', 'CRM'],
        'Cigars and Tobacco': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Marketing Automation'],
        'Oil & Gas Companies': ['Website', 'CRM', 'Inventory Mgmt', 'Analytics & Reporting', 'Security Systems', 'Cloud Backup'],
        'Government Services': ['Website', 'CRM', 'Cloud Backup', 'Analytics & Reporting', 'Security Systems', 'Communication'],
        'Driving Schools': ['Website', 'Booking System', 'Learning Mgmt', 'CRM', 'SMS/Notification', 'Communication'],
        'Trading Platform': ['Website', 'CRM', 'Security Systems', 'Analytics & Reporting', 'Marketing Automation', 'Cloud Backup'],
        'Optical Shop': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'CRM', 'Marketing Automation'],
        'Marketing': ['Website', 'Social Media Mgmt', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication'],
        '3D Technology': ['Website', 'CRM', 'Cloud Backup', 'Marketing Automation', 'Analytics & Reporting', 'Communication'],
        'B2B': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Auctioneers': ['Website', 'E-commerce', 'CRM', 'Marketing Automation', 'SMS/Notification', 'Analytics & Reporting'],
        'Property Consultants': ['Website', 'CRM', 'Marketing Automation', 'Analytics & Reporting', 'Communication', 'Social Media Mgmt'],
        'Computers Hardware': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Delivery Logistics', 'CRM'],
        'Holiday Homes': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup'],
        'Auctions': ['Website', 'E-commerce', 'CRM', 'Marketing Automation', 'SMS/Notification', 'Analytics & Reporting'],
        'Computer Training': ['Website', 'Learning Mgmt', 'CRM', 'Marketing Automation', 'Cloud Backup', 'SMS/Notification'],
        'Leasing': ['Website', 'CRM', 'Inventory Mgmt', 'Accounting/ERP', 'Communication', 'Analytics & Reporting'],
        'Funeral Directors': ['Website', 'CRM', 'SMS/Notification', 'Communication', 'Marketing Automation', 'Cloud Backup'],
        'Religion': ['Website', 'Communication', 'Social Media Mgmt', 'Cloud Backup', 'Marketing Automation', 'Analytics & Reporting'],
        'Fitness': ['Website', 'Booking System', 'CRM', 'POS System', 'Marketing Automation', 'Social Media Mgmt'],
        'Voluntary Organisations': ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication', 'Analytics & Reporting'],
        'Factories': ['Website', 'Inventory Mgmt', 'Accounting/ERP', 'Security Systems', 'Analytics & Reporting', 'Cloud Backup'],
        'Recycling': ['Website', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Analytics & Reporting', 'Communication'],
        'Fire Safety Equipment': ['Website', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Communication', 'Marketing Automation'],
        'Wine and Beer': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'CRM', 'Delivery Logistics'],
        'Product Development': ['Website', 'CRM', 'Cloud Backup', 'Analytics & Reporting', 'Marketing Automation', 'Communication'],
        'Visa Agencies': ['Website', 'CRM', 'Booking System', 'Marketing Automation', 'SMS/Notification', 'Communication'],
        'Construction Services': ['Website', 'CRM', 'Accounting/ERP', 'Cloud Backup', 'Communication', 'Analytics & Reporting'],
        'Gifts': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'Delivery Logistics'],
        'Mental Health Care': ['Website', 'Patient Mgmt', 'Booking System', 'CRM', 'SMS/Notification', 'Cloud Backup'],
        'Industrial Automation': ['Website', 'CRM', 'Analytics & Reporting', 'Cloud Backup', 'Security Systems', 'Communication'],
        'Airports': ['Website', 'Booking System', 'CRM', 'Security Systems', 'Analytics & Reporting', 'Communication'],
        'Fishing Equipment': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'Delivery Logistics', 'CRM'],
        'Textile': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Accounting/ERP'],
        'E-Learning': ['Website', 'Learning Mgmt', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication'],
        'Furniture Manufacturers': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Marketing Automation'],
        'Casinos': ['Website', 'POS System', 'CRM', 'Security Systems', 'Analytics & Reporting', 'Marketing Automation'],
        'Furniture': ['Website', 'E-commerce', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Social Media Mgmt'],
        'Haulage': ['Website', 'Delivery Logistics', 'CRM', 'Analytics & Reporting', 'Communication', 'Cloud Backup'],
        'Science': ['Website', 'CRM', 'Cloud Backup', 'Analytics & Reporting', 'Communication', 'Marketing Automation'],
        'Nursing and Care': ['Website', 'Patient Mgmt', 'Booking System', 'CRM', 'SMS/Notification', 'Cloud Backup'],
        'Letting Agents': ['Website', 'CRM', 'Marketing Automation', 'Communication', 'Cloud Backup', 'Accounting/ERP'],
        'Tourism': ['Website', 'Booking System', 'CRM', 'Marketing Automation', 'Social Media Mgmt', 'Analytics & Reporting'],
        'Finances & Insurance': ['Website', 'CRM', 'Accounting/ERP', 'Security Systems', 'Analytics & Reporting', 'Cloud Backup'],
        'Pest Control': ['Website', 'Booking System', 'CRM', 'Delivery Logistics', 'SMS/Notification', 'Marketing Automation'],
        'Opticians': ['Website', 'E-commerce', 'Inventory Mgmt', 'POS System', 'CRM', 'Booking System'],
        'Software': ['Website', 'Cloud Backup', 'CRM', 'Analytics & Reporting', 'Marketing Automation', 'Security Systems'],
        'Animal Shelters': ['Website', 'CRM', 'Social Media Mgmt', 'Cloud Backup', 'Communication', 'Marketing Automation'],
        'Chemicals': ['Website', 'Inventory Mgmt', 'CRM', 'Delivery Logistics', 'Analytics & Reporting', 'Cloud Backup'],
        'Storage Services': ['Website', 'CRM', 'Inventory Mgmt', 'Marketing Automation', 'Communication', 'Cloud Backup'],
        'Laundry and Dry Cleaning': ['Website', 'Online Ordering', 'POS System', 'Inventory Mgmt', 'Delivery Logistics', 'SMS/Notification'],
        'Social Work Services': ['Website', 'CRM', 'Cloud Backup', 'Communication', 'Analytics & Reporting', 'Marketing Automation'],
        'Tailors and Alterations': ['Website', 'Social Media Mgmt', 'E-commerce', 'POS System', 'Inventory Mgmt', 'Marketing Automation'],
        'Warehousing': ['Website', 'Inventory Mgmt', 'Delivery Logistics', 'CRM', 'Analytics & Reporting', 'Cloud Backup'],
        'Infrastructure': ['Website', 'CRM', 'Analytics & Reporting', 'Cloud Backup', 'Communication', 'Security Systems'],
        'Kids': ['Website', 'E-commerce', 'Social Media Mgmt', 'Inventory Mgmt', 'Marketing Automation', 'CRM'],
        'Vehicle Manufacturers': ['Website', 'Inventory Mgmt', 'CRM', 'Accounting/ERP', 'Analytics & Reporting', 'Communication'],
        'Bookmakers': ['Website', 'CRM', 'Analytics & Reporting', 'Communication', 'Marketing Automation', 'Cloud Backup'],
        'Nanny Agency': ['Website', 'CRM', 'Booking System', 'Marketing Automation', 'SMS/Notification', 'Cloud Backup'],
    };

    // Get services for a category (with fallback)
    function getServices(cat) {
        if (industryNeeds[cat]) return industryNeeds[cat];
        return ['Website', 'CRM', 'Marketing Automation', 'Cloud Backup', 'Communication', 'Social Media Mgmt'];
    }

    // ======== Pre-processing ========
    let categories = {};
    for (const b of businesses) {
        categories[b.category || 'Uncategorized'] = (categories[b.category || 'Uncategorized'] || 0) + 1;
    }
    const sortedCats = Object.entries(categories).sort((a,b) => b[1] - a[1]);

    // ======== SHEET 1: DASHBOARD ========
    const wsDash = wb.addWorksheet('Dashboard');
    wsDash.columns = [{ header: '', key: 'l', width: 35 }, { header: '', key: 'v', width: 30 }];
    
    wsDash.mergeCells('A1:B1');
    wsDash.getCell('A1').value = 'BUSINESS ASSESSMENT & AUTOMATION SYSTEM';
    wsDash.getCell('A1').style = { font: { bold: true, size: 16, color: { argb: '1565C0' } }, alignment: { horizontal: 'center' } };
    
    wsDash.mergeCells('A3:B3');
    wsDash.getCell('A3').value = 'SERVICE PORTFOLIO & PRICING';
    wsDash.getCell('A3').style = { font: { bold: true, size: 13, color: { argb: '1565C0' } } };
    
    let row = 4;
    for (const [svc, info] of Object.entries(serviceCatalog)) {
        wsDash.getCell('A' + row).value = svc;
        wsDash.getCell('B' + row).value = '$' + info.base + ' - ' + info.desc;
        wsDash.getCell('A' + row).style = { font: { bold: true } };
        row++;
    }

    row += 2;
    wsDash.getCell('A' + row).value = 'TOTAL SERVICEABLE MARKET';
    wsDash.getCell('A' + row).style = { font: { bold: true, size: 13, color: { argb: '1565C0' } } };
    row++;
    wsDash.getCell('A' + row).value = 'Total Businesses';
    wsDash.getCell('B' + row).value = businesses.length;
    row++;
    wsDash.getCell('A' + row).value = 'Categories Covered';
    wsDash.getCell('B' + row).value = sortedCats.length;
    row++;
    wsDash.getCell('A' + row).value = 'Services Offered';
    wsDash.getCell('B' + row).value = Object.keys(serviceCatalog).length;

    // ======== SHEET 2: COMPANY ASSESSMENTS ========
    const wsAssess = wb.addWorksheet('Company Assessments');
    const assessCols = [
        { header: '#', key: 'id', width: 4 },
        { header: 'Business Name', key: 'name', width: 35 },
        { header: 'Category', key: 'cat', width: 22 },
        { header: 'Current Website', key: 'web', width: 14 },
        { header: 'Digital Maturity', key: 'maturity', width: 14 },
        { header: 'Total Est. Value', key: 'total', width: 12 },
    ];
    
    // Add 6 service columns
    for (let i = 1; i <= 6; i++) {
        assessCols.push({ header: 'Service ' + i, key: 'svc' + i, width: 20 });
        assessCols.push({ header: 'Est. $', key: 'val' + i, width: 8 });
    }
    assessCols.push({ header: 'Sales Stage', key: 'stage', width: 14 });
    assessCols.push({ header: 'Notes', key: 'notes', width: 35 });
    
    wsAssess.columns = assessCols;
    wsAssess.getRow(1).style = hdr;
    wsAssess.getRow(1).height = 30;
    
    let assessId = 1;
    for (const b of businesses) {
        const website = b.verifiedWebsite || b.website || '';
        const services = getServices(b.category || '');
        
        let maturity = 'None';
        let baseValue = 0;
        if (website && b.websiteWorking === true) { maturity = 'Basic'; baseValue = 1; }
        else if (website) { maturity = 'Weak'; baseValue = 0; }
        
        const hasEmail = b.email && b.email.length > 0;
        const hasPhone = (b.phone && b.phone.length > 0) || (b.mobile && b.mobile.length > 0);
        if (hasEmail && hasPhone && maturity === 'Basic') maturity = 'Moderate';
        if (maturity === 'Moderate' && website) maturity = 'Good';
        
        const rowData = { id: assessId++, name: b.name, cat: b.category || '', web: website ? (b.websiteWorking === true ? 'Working' : 'Broken') : 'None', maturity: maturity, total: 0 };
        
        let totalVal = 0;
        for (let i = 0; i < 6; i++) {
            if (services[i]) {
                const svcInfo = serviceCatalog[services[i]] || { base: 300, desc: '' };
                rowData['svc' + (i+1)] = services[i];
                rowData['val' + (i+1)] = svcInfo.base;
                totalVal += svcInfo.base;
            } else {
                rowData['svc' + (i+1)] = '';
                rowData['val' + (i+1)] = 0;
            }
        }
        
        rowData.total = totalVal;
        rowData.stage = 'New';
        rowData.notes = '';
        
        wsAssess.addRow(rowData);
    }

    // ======== SHEET 3: ASSESSMENT SUMMARY ========
    const wsSummary = wb.addWorksheet('Assessment Summary');
    wsSummary.columns = [
        { header: 'Category', key: 'cat', width: 28 },
        { header: 'Total', key: 'count', width: 8 },
        { header: 'Avg Est. Value', key: 'avg', width: 14 },
        { header: 'Total Portfolio Value', key: 'portfolio', width: 20 },
        { header: 'Top Service 1', key: 's1', width: 22 },
        { header: 'Top Service 2', key: 's2', width: 22 },
        { header: 'Top Service 3', key: 's3', width: 22 },
    ];
    wsSummary.getRow(1).style = hdr;
    wsSummary.getRow(1).height = 25;
    
    for (const [cat, count] of sortedCats) {
        const services = getServices(cat);
        wsSummary.addRow({
            cat: cat,
            count: count,
            avg: '$' + (services.reduce((sum, s) => sum + (serviceCatalog[s] ? serviceCatalog[s].base : 300), 0) / services.length).toFixed(0),
            portfolio: '$' + (count * services.reduce((sum, s) => sum + (serviceCatalog[s] ? serviceCatalog[s].base : 300), 0)).toLocaleString(),
            s1: services[0] || '',
            s2: services[1] || '',
            s3: services[2] || '',
        });
    }

    // ======== SHEET 4: SERVICE CATALOG ========
    const wsCat2 = wb.addWorksheet('Service Catalog');
    wsCat2.columns = [
        { header: 'Service', key: 'svc', width: 25 },
        { header: 'Base Price', key: 'price', width: 12 },
        { header: 'Description', key: 'desc', width: 70 },
        { header: 'Target Industries', key: 'industries', width: 80 },
    ];
    wsCat2.getRow(1).style = hdr;
    wsCat2.getRow(1).height = 25;
    
    for (const [svc, info] of Object.entries(serviceCatalog)) {
        let targets = [];
        for (const [cat, needs] of Object.entries(industryNeeds)) {
            if (needs.includes(svc)) targets.push(cat);
        }
        wsCat2.addRow({ svc: svc, price: '$' + info.base, desc: info.desc, industries: targets.slice(0, 10).join(', ') + (targets.length > 10 ? '...' : '') });
    }

    // ======== SHEET 5: TOP 20 CATEGORY ASSESSMENTS ========
    for (const [cat, count] of sortedCats.slice(0, 20)) {
        let safeName = cat.replace(/[\\/:*?"<>|]/g, ' ').substring(0, 31).trim();
        if (!safeName) safeName = 'Other';
        
        const wsCA = wb.addWorksheet(safeName);
        wsCA.columns = [
            { header: '#', key: 'id', width: 4 },
            { header: 'Business Name', key: 'name', width: 38 },
            { header: 'Current Site', key: 'site', width: 10 },
            { header: 'Services Needed', key: 'needs', width: 55 },
            { header: 'Total Est.', key: 'total', width: 10 },
            { header: 'Contact', key: 'contact', width: 35 },
            { header: 'Stage', key: 'stage', width: 12 },
        ];
        wsCA.getRow(1).style = hdr;
        wsCA.getRow(1).height = 25;
        
        const services = getServices(cat);
        let caId = 1;
        for (const b of businesses) {
            if ((b.category || 'Uncategorized') === cat) {
                const website = b.verifiedWebsite || b.website || '';
                const total = services.reduce((sum, s) => sum + (serviceCatalog[s] ? serviceCatalog[s].base : 300), 0);
                wsCA.addRow({
                    id: caId++,
                    name: b.name,
                    site: website ? (b.websiteWorking === true ? 'Yes' : 'Broken') : 'No',
                    needs: services.join(', '),
                    total: '$' + total,
                    contact: b.email || b.phone || b.mobile || '',
                    stage: 'New'
                });
            }
        }
    }

    await wb.xlsx.writeFile('Kigali_Business_Assessment.xlsx');
    console.log('ASSESSMENT SYSTEM COMPLETE!');
    console.log('File: Kigali_Business_Assessment.xlsx');
    console.log('');
    console.log('Sheets:');
    console.log('  1. Dashboard - Service portfolio with pricing');
    console.log('  2. Company Assessments - Every business assessed with 6 recommended services + values');
    console.log('  3. Assessment Summary - Per-category breakdown with portfolio value');
    console.log('  4. Service Catalog - Which industries need each service');
    console.log('  5. Top 20 Category sheets - Targeted industry assessment lists');
    console.log('');
    console.log('20 Services Available:');
    Object.entries(serviceCatalog).forEach(([s, i]) => console.log('  - ' + s + ' ($' + i.base + '): ' + i.desc));
})();