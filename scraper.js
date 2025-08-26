const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const axios = require('axios');
const fs = require('fs');
const xlsx = require('xlsx');

// Add stealth plugin to avoid bot detection
puppeteer.use(StealthPlugin());

// === 1. CONFIGURATION ===
const VERIFICATION_URL = "https://trailhead.salesforce.com/en/credentials/verification/";
const SITE_KEY = "6LfqUVUUAAAAAFKKuRyeZRh0iiBjVcdhjWIThGSV"; // Site key for the reCAPTCHA
const API_KEY = "c655334f072da9b94b0cf835225e2c5b"; // <-- 🚨 PASTE YOUR 2CAPTCHA API KEY HERE
const MAX_RETRIES = 3;
const EXCEL_FILE_PATH = '/data/input1.xlsx' // The file to read from and write to

/**
 * Delays execution for a specified amount of time.
 */
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Solves the reCAPTCHA using the 2Captcha service.
 */
async function solveRecaptcha(siteKey, pageUrl) {
    if (API_KEY === "YOUR_2CAPTCHA_API_KEY") {
        throw new Error("Please update the API_KEY variable with your 2Captcha API key.");
    }
    console.log("Submitting CAPTCHA to 2Captcha...");
    const submitUrl = `http://2captcha.com/in.php?key=${API_KEY}&method=userrecaptcha&googlekey=${siteKey}&pageurl=${pageUrl}&json=1`;
    const response = await axios.get(submitUrl);
    if (response.data.status !== 1) throw new Error(`Captcha submit error: ${response.data.request || 'Unknown'}`);
    
    const captchaId = response.data.request;
    console.log(`Waiting for CAPTCHA to be solved (ID: ${captchaId})...`);
    const resultUrl = `http://2captcha.com/res.php?key=${API_KEY}&action=get&id=${captchaId}&json=1`;

    for (let i = 0; i < 40; i++) {
        await sleep(5000 + Math.random() * 1500);
        const resultResponse = await axios.get(resultUrl);
        if (resultResponse.data.status === 1) {
            console.log("✅ CAPTCHA solved.");
            return resultResponse.data.request;
        }
        if (resultResponse.data.request !== 'CAPCHA_NOT_READY') {
            throw new Error(`2Captcha solve error: ${resultResponse.data.request}`);
        }
    }
    throw new Error("❌ CAPTCHA was not solved in time.");
}

/**
 * Scrapes the full credential names for a given email from the verification website.
 */
async function getCredentialsForEmail(browser, email) {
    let page = null;
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
            console.log(`\n--- Processing ${email} (Attempt ${attempt}/${MAX_RETRIES}) ---`);
            page = await browser.newPage();
            await page.setViewport({ width: 1280, height: 800 });
            await page.goto(VERIFICATION_URL, { waitUntil: 'networkidle2' });

            await page.waitForSelector("xpath///input[@placeholder='Full Name or Email']");
            await page.type("xpath///input[@placeholder='Full Name or Email']", email, { delay: 50 });
            await page.click("xpath///button[text()='Search']");

            await page.waitForSelector('iframe[src*="recaptcha"]', { visible: true, timeout: 10000 });
            const token = await solveRecaptcha(SITE_KEY, VERIFICATION_URL);
            await page.evaluate((token) => {
                document.getElementById('g-recaptcha-response').value = token;
                const container = document.querySelector('.g-recaptcha');
                const callback = container?.getAttribute('data-callback');
                if (callback && typeof window[callback] === 'function') window[callback](token);
            }, token);

            const credentialButtonXPath = "//button[contains(text(), 'Credentials info')]";
            try {
                // === FIX: Increased timeout to give the page more time to load results ===
                await page.waitForSelector(`xpath/${credentialButtonXPath}`, { timeout: 30000 });
                console.log("✅ Credentials found. Scraping details...");
                await page.click(`xpath/${credentialButtonXPath}`);
                
                const printViewXPath = "//a[contains(text(), 'Print View')]";
                await page.waitForSelector(`xpath/${printViewXPath}`, { visible: true });
                
                const pageTarget = page.target();
                await page.click(`xpath/${printViewXPath}`);
                const newTarget = await browser.waitForTarget(target => target.opener() === pageTarget, { timeout: 10000 });
                const newPage = await newTarget.page();

                if (!newPage) throw new Error("Failed to capture the new 'Print View' tab.");

                await newPage.waitForSelector('.verification-overlay__items-container', { timeout: 15000 });
                const credentialsData = await newPage.evaluate(() => 
                    Array.from(document.querySelectorAll('.verification-overlay__items-container > div'))
                         .map(row => row.querySelector('.slds-col.slds-p-horizontal--medium')?.innerText.trim())
                         .filter(Boolean)
                );
                
                await newPage.close();
                await page.close();
                return credentialsData;

            } catch (error) {
                if (error.name === 'TimeoutError') {
                    console.log(`No credentials found for ${email} (button did not appear in time).`);
                    await page.close();
                    return [];
                }
                // === FIX: Better error logging for other potential issues ===
                console.error(`An error occurred while trying to scrape details for ${email}: ${error.message}`);
                throw error;
            }
        } catch (error) {
            console.error(`\n❌ Error on attempt ${attempt} for ${email}: ${error.message}`);
            if (page) await page.close();
            if (attempt === MAX_RETRIES) return null;
        }
    }
    return null;
}

/**
 * Main orchestrator function to dynamically build the Excel file.
 */
async function run() {
    if (!fs.existsSync(EXCEL_FILE_PATH)) {
        console.error(`Error: The file "${EXCEL_FILE_PATH}" was not found.`);
        return;
    }
    
    // 1. Read the Excel file into an array of objects
    console.log(`Reading data from "${EXCEL_FILE_PATH}"...`);
    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    const emailsToProcess = data.filter(row => row.Email);
    if (emailsToProcess.length === 0) {
        console.log("No rows with an 'Email' found in the file. Exiting.");
        return;
    }
    
    // Keep track of all certification columns we've created
    const allCertHeaders = new Set();

    console.log(`Starting to process ${emailsToProcess.length} email(s)...`);
    const browser = await puppeteer.launch({ 
        headless: true,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-background-timer-throttling'
        ]
    });

    // 2. Loop through each person (row) to scrape and update data in memory
    for (const row of emailsToProcess) {
        const scrapedNames = await getCredentialsForEmail(browser, row.Email);

        if (scrapedNames === null) {
            row["Didn't Find"] = "Scraping Failed";
            continue;
        }

        row["No. Of certifications"] = scrapedNames.length;
        row["Didn't Find"] = scrapedNames.length > 0 ? "Found" : "NA";
        
        const foundCertsSet = new Set(scrapedNames);

        // For each certification this person has...
        for (const certName of foundCertsSet) {
            // Add it to our master list of headers
            allCertHeaders.add(certName);
            // Mark this person as having it
            row[certName] = "Yes";
        }
        console.log(`✅ Processed data for ${row.Email}.`);
    }

    await browser.close();

    // 3. Post-processing: Ensure all rows have all columns
    for (const row of data) {
        for (const header of allCertHeaders) {
            // If a row is missing a certification column that was added for someone else,
            // fill it with a placeholder.
            if (!row.hasOwnProperty(header)) {
                row[header] = "-";
            }
        }
    }

    // 4. Write the completely updated data back to the file
    console.log(`\n🎉 All processing complete. Writing updated data back to "${EXCEL_FILE_PATH}"...`);
    
    const updatedSheet = xlsx.utils.json_to_sheet(data);
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, updatedSheet, sheetName);
    xlsx.writeFile(newWorkbook, EXCEL_FILE_PATH);

    console.log(`✅ Success! The file "${EXCEL_FILE_PATH}" has been updated with dynamic columns.`);
}

// Start the process
run().catch(err => {
    console.error("\n💥 A critical error occurred in the main process:", err);
    process.exit(1);
});
