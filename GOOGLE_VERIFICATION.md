# Google Maps Business Verification System

## Setup Instructions

### 1. Get a Google API Key
1. Go to https://console.cloud.google.com/
2. Create a new project (or select existing)
3. Enable these APIs:
   - **Places API** (for Google Maps business data)
   - **Custom Search API** (for fallback website lookup)
4. Go to Credentials → Create API Key
5. Copy the key and paste it into `config.json`

### 2. Install Dependencies
```bash
cd business-scanner
npm install googleapis axios
```

### 3. Configure
Edit `config.json` with your API key.

### 4. How It Works

The system uses Google Places API Text Search to find each business on Google Maps, which returns:
- ✅ Real website URL
- ✅ Verified phone number
- ✅ Full formatted address  
- ✅ Google rating & reviews
- ✅ Place ID (for cross-reference)

For businesses not found on Maps, it falls back to Google Custom Search.

---

## Files Created

| File | Purpose |
|---|---|
| `config.json` | API key configuration |
| `google_verify.js` | Main verification script |
| `verify-progress.json` | Auto-saves progress every 50 businesses |
| `verification-dashboard.html` | Web dashboard to track & review results |
| `Kigali_Verified.xlsx` | Final verified output |
