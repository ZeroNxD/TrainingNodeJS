const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const express = require('express')
const app = express();
const PORT = process.env.PORT || 3000;

// Cara langsung Upload Filenya langsung hit ke APInya
app.get('/', (req, res) => {
    res.send("Server is running! The script is already executed.");
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
})

const filePath = "D:/Latihan UNIAIR/ProjectKe2-RPA/ProjectRPAUNIAIR/DataUser.xlsx";

// Make Sure kalau Path ke Datasetnya ada
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
        const instance = axios.create({
            baseURL: 'http://127.0.0.1:8000',
            withCredentials: true,
            headers: { 'Content-Type': 'application/json' }
        });
        
    
        // Setelah di read, coba kita send datanya ke Laravel via API.
        // Masih make API ini coba diarahin ke Puppeteer.
        instance.post('/api/upload', { listdata: NewData })
            .then(response => {
                console.log('Data sent to Laravel Project Successfully : ', response.data);
                // Lakuin proses updating kalau semisal udah diupload.
                NewData.array.forEach(NewUser => {
                    const index = Data.findIndex(user => user.Number === NewUser.Number);
                    if (index !== -1){
                        Data[index].Status = 'Success';
                    }
                });

                // Habis Update Value, tinggal update ke excel file
                const NewWorkSheet = xlsx.utils.json_to_sheet(Data);
                WorkBook.Sheets[WorkBook.SheetNames[0]] == NewWorkSheet;
                xlsx.writeFile(WorkBook, filePath);
                console.log("Excel File Updated!!")
            }).catch(error => {
                console.error("Error Sending Data: ", error.response ? error.response.data : error.message);
            })
    } else {
        console.log("No Data to send, make new data at DataUser.xlsx")
    }

} else {
    console.error("File not found:", filePath);
}
