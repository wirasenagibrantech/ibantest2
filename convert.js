// convert.js
const fs = require("fs");
const xlsx = require("xlsx");

const file = xlsx.readFile("data.xlsx");
const sheet = file.Sheets[file.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(sheet, { raw: false });

// Pastikan folder output exist
if (!fs.existsSync("data")) fs.mkdirSync("data");

// Convert Excel number / date to real Date()
function parseDate(d) {
    return new Date(d);
}

// --- Prepare data ---
let fullData = [];
let rekapMingguan = {};
let rekapBulanan = {};
let rekapTahunan = {};

rows.forEach((row) => {
    const date = parseDate(row.tanggal || row.Tanggal);
    const penjualan = Number(row.penjualan || row.Penjualan || 0);

    if (isNaN(date.getTime())) return; // skip kalau tanggal invalid

    const year = date.getFullYear();
    const month = date.getMonth() + 1;        // 1–12
    const week = `${year}-W${String(
        Math.ceil((date.getDate() + new Date(year, 0, 1).getDay()) / 7)
    ).padStart(2, "0")}`;

    // 1. Simpan data full
    fullData.push({
        tanggal: date.toISOString().split("T")[0],
        year,
        month,
        week,
        penjualan
    });

    // 2. Rekap mingguan
    if (!rekapMingguan[week]) rekapMingguan[week] = 0;
    rekapMingguan[week] += penjualan;

    // 3. Rekap bulanan
    const bulanKey = `${year}-${String(month).padStart(2, "0")}`;
    if (!rekapBulanan[bulanKey]) rekapBulanan[bulanKey] = 0;
    rekapBulanan[bulanKey] += penjualan;

    // 4. Rekap tahunan
    if (!rekapTahunan[year]) rekapTahunan[year] = 0;
    rekapTahunan[year] += penjualan;
});

// Save outputs
fs.writeFileSync("data/data.json", JSON.stringify(fullData, null, 2));
fs.writeFileSync("data/rekap-mingguan.json", JSON.stringify(rekapMingguan, null, 2));
fs.writeFileSync("data/rekap-bulanan.json", JSON.stringify(rekapBulanan, null, 2));
fs.writeFileSync("data/rekap-tahunan.json", JSON.stringify(rekapTahunan, null, 2));

console.log("Konversi selesai → folder /data sudah terupdate");
