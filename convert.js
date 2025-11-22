// convert.js
const XLSX = require("xlsx");
const fs = require("fs");

// === BACA FILE EXCEL ===
const workbook = XLSX.readFile("data.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

const bulanIndo = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];

// --- KONVERSI TANGGAL EXCEL ---
function excelToJSDate(excel) {
  return new Date((excel - 25569) * 86400 * 1000);
}

// === DATA NORMAL ===
const data = rows.map(r => {
  const d = excelToJSDate(r.tanggal);
  return {
    tanggal: `${d.getDate()} ${bulanIndo[d.getMonth()]} ${d.getFullYear()}`,
    tahun: d.getFullYear(),
    bulan: d.getMonth() + 1,
    minggu: Math.ceil((d.getDate()) / 7),
    penjualan: r.penjualan
  };
});

// === REKAP MINGGUAN ===
const rekapMingguan = {};
data.forEach(r => {
  const key = `${r.tahun}-${r.bulan}-M${r.minggu}`;
  if (!rekapMingguan[key]) rekapMingguan[key] = 0;
  rekapMingguan[key] += r.penjualan;
});

// === REKAP BULANAN ===
const rekapBulanan = {};
data.forEach(r => {
  const key = `${r.tahun}-${r.bulan}`;
  if (!rekapBulanan[key]) rekapBulanan[key] = 0;
  rekapBulanan[key] += r.penjualan;
});

// === REKAP TAHUNAN ===
const rekapTahunan = {};
data.forEach(r => {
  const key = `${r.tahun}`;
  if (!rekapTahunan[key]) rekapTahunan[key] = 0;
  rekapTahunan[key] += r.penjualan;
});

// === SIMPAN SEMUA FILE JSON ===
if (!fs.existsSync("json")) fs.mkdirSync("json");

fs.writeFileSync("json/data.json", JSON.stringify(data, null, 2));
fs.writeFileSync("json/rekap-mingguan.json", JSON.stringify(rekapMingguan, null, 2));
fs.writeFileSync("json/rekap-bulanan.json", JSON.stringify(rekapBulanan, null, 2));
fs.writeFileSync("json/rekap-tahunan.json", JSON.stringify(rekapTahunan, null, 2));

console.log("✔ Semua JSON berhasil dibuat!");
