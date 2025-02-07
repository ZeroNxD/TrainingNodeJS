const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const express = require('express')
const app = express();
const PORT = process.env.PORT || 3000;


app.get('/', (req, res) => {
    res.send("Server is running! The script is already executed.");
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
})

async function updateStatusInLaravel(email, status) {
    try {
        const response = await axios.post('http://127.0.0.1:8000/update-status', {
            email: email,
            status: status
        });
        console.log(`Status updated in Laravel for ${email}: ${response.data.message}`);
    } catch (error) {
        console.error(`Failed to update status in Laravel for ${email}:`, error.message);
    }
}

const filePath = "./DataUser.xlsx";

// Make Sure kalau Path ke Datasetnya ada
(async () => {
    if (fs.existsSync(filePath)) {
        // Coba Read Excel File dulu
        const WorkBook = xlsx.readFile(filePath);
        console.log("File read successfully.");
        const WorkSheet = WorkBook.Sheets[WorkBook.SheetNames[0]];
        const Data = xlsx.utils.sheet_to_json(WorkSheet, {header:0}); // Diubah ke JSON dulu biar bisa dapat konteks Excelnya
        // Cek Datanya dah kebaca atau belum
        console.log("Data Read Successfully!! -> ", Data);
    
        // Kita lakukan filterisasi dimana yang kita ambil hanya yang "Failure" dan "NULL"
        const NewData = Data.filter(user => !user.Status || user.Status === "Failure");
        console.log("List Data Baru -> ", NewData);
    
        if(NewData.length > 0){
            (async () => {
                const browser = await puppeteer.launch({headless: false});
                const page = await browser.newPage();
                await page.goto("http://127.0.0.1:8000/")
    
                for (let i = 0; i < NewData.length; i++){
                    const user = NewData[i];
                    try {
                        console.log(`Processing User: ${user.Name}, ${user.Email}`);
                        await page.click(".btn")
                        await page.waitForSelector('.mb-3 > input:nth-child(2)');
                        await page.waitForSelector('.mb-5 > input:nth-child(2)');
                        await page.type(".mb-3 > input:nth-child(2)", user.Name, {delay:200});
                        await page.type(".mb-5 > input:nth-child(2)", user.Email, {delay:200});
                        await Promise.all([
                            await page.click('button[type="submit"]'),
                            await page.waitForNavigation({ waitUntil: 'networkidle2' }),
                        ]);
                        
                        const success = await page.$('.alert-success');
                        console.log(success);
    
                        if (success) {
                            console.log(`Registration successful for ${user.Name}`);
                            user.Status = 'Success';
                            await updateStatusInLaravel(user.Email, 'Success')
                        } else {
                            console.log(`Registration failed for ${user.Name}`);
                            user.Status = 'Failure';
                            await updateStatusInLaravel(user.Email, 'Failure')
                        }
                    } catch (error){
                        console.error(`Error processing user ${user.Name}:`, error);
                        user.Status = 'Failure';
                    }
                }
                await page.reload();
                await browser.close();
    
                const NewWorkSheet = xlsx.utils.json_to_sheet(Data);
                WorkBook.Sheets[WorkBook.SheetNames[0]] = NewWorkSheet;
                xlsx.writeFile(WorkBook, filePath);
                console.log("Excel File Updated!!");
            })();
        } else {
            console.log("No Data to send, make new data at DataUser.xlsx")
        }
    
    } else {
        console.error("File not found:", filePath);
    }
    
})();

