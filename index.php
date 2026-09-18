<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Generator Surat Keterangan LSP LAS</title>

<!-- Library PDF: langsung canvas -> jsPDF, sehingga selalu 1 halaman A4 -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<style>
*{box-sizing:border-box}

body{
    margin:0;
    padding:20px;
    background:#f3f3f3;
    font-family:Arial,sans-serif;
    color:#111;
}

.container{
    max-width:1300px;
    margin:auto;
    background:#fff;
    padding:25px;
    border-radius:10px;
    box-shadow:0 0 10px rgba(0,0,0,.1);
}

.page-title{
    text-align:center;
    margin:0 0 22px;
}

.form-group{margin-bottom:14px}

label{
    display:block;
    font-weight:bold;
    margin-bottom:5px;
}

input,select{
    width:100%;
    padding:10px;
    border:1px solid #ccc;
    border-radius:5px;
    font-size:14px;
    background:#fff;
}

button{
    padding:10px 18px;
    border:0;
    border-radius:5px;
    cursor:pointer;
    color:#fff;
    margin:5px 8px 10px 0;
    font-weight:bold;
}

button:active{transform:translateY(1px)}
.btn-generate{background:#2563eb}
.btn-pdf{background:#16a34a}
.btn-excel{background:#f59e0b}
.btn-import{background:#7c3aed}

.preview-wrapper{
    margin-top:30px;
    padding:20px;
    background:#ddd;
    overflow:auto;
}

/* ================================
   SURAT A4 - 794 x 1123 px
   Isi surat dibuat seimbang kiri-kanan
================================ */
.preview{
    width:794px;
    height:1123px;
    background:#fff;
    margin:0 auto;
    padding:30px 62px 78px;
    box-sizing:border-box;
    font-family:"Times New Roman",Times,serif;
    font-size:15px;
    line-height:1.45;
    position:relative;
    overflow:hidden;
    color:#000;
}

/* Clone PDF tidak boleh terpengaruh wrapper */
.pdf-clone{
    width:794px !important;
    height:1123px !important;
    margin:0 !important;
    padding:30px 62px 78px !important;
    position:absolute !important;
    left:0 !important;
    top:0 !important;
    transform:none !important;
    zoom:1 !important;
    overflow:hidden !important;
    background:#fff !important;
}

/* ================================
   KOP
================================ */
.kop{
    width:100%;
    margin:0;
    padding:0 0 6px;
    border-bottom:2px solid #000;
}

.kop table{
    width:100%;
    border-collapse:collapse;
    margin:0;
}

.kop td{
    border:0;
    padding:0;
    vertical-align:middle;
}

.kop .logo-cell-left{width:14%;text-align:left}
.kop .logo-cell-center{width:72%;text-align:center}
.kop .logo-cell-right{width:14%;text-align:right}

.logo-lsp{
    width:62px;
    height:auto;
    display:block;
}

.logo-bnsp{
    width:92px;
    height:auto;
    display:block;
    margin-left:auto;
}

.nama-lembaga{
    text-align:center;
    line-height:1.05;
}

.nama-lembaga h1{
    margin:0;
    font-size:20px;
    font-weight:bold;
    white-space:nowrap;
}

.nama-lembaga h2{
    margin:3px 0 4px;
    font-size:16px;
    font-weight:bold;
}

.izin{
    font-size:7.5px;
    line-height:1.25;
    white-space:nowrap;
}

/* ================================
   JUDUL
================================ */
.judul{
    text-align:center;
    margin-top:5px;
    margin-bottom:10px;
    line-height:1;
}

.judul h2{
    margin:0;
    padding:0;
    font-size:17px;
    line-height:1;
    text-decoration:underline;
    font-weight:bold;
}

.judul h3{
    margin:2px 0 0;
    padding:0;
    font-size:13px;
    line-height:1;
    font-weight:normal;
}

/* ================================
   ISI
================================ */
.isi{
    text-align:justify;
    margin-top:0;
}

.nama-asesi{
    text-align:center;
    font-size:18px;
    font-weight:bold;
    margin:12px 0;
}

.skema{
    text-align:center;
    font-size:17px;
    font-weight:bold;
    margin:12px 0;
}

.hasil{
    font-size:17px;
    font-weight:bold;
    text-align:center;
    margin:3px 0 12px;
}

/* ================================
   TANDA TANGAN
================================ */
.ttd{
    margin-top:20px;
    height:145px;
    position:relative;
}

.ttd-area{
    width:285px;
    margin-left:auto;
    position:relative;
    text-align:center;
}

.ttd-kota{
    font-size:15px;
    margin:0 0 2px 0;
    line-height:1.2;
}

.ttd-jabatan{
    font-size:15px;
    font-weight:bold;
    margin:0;
    line-height:1.2;
}

/* Area tanda tangan + cap dibuat lebih besar.
   Cap sengaja menempel/overlap sekitar setengah dengan tanda tangan,
   tetapi keduanya tetap berada di atas nama direktur. */
.signature-visual{
    width:225px;
    height:88px;
    margin:3px auto 0;
    position:relative;
}

.ttd-img{
    position:absolute;
    width:130px;
    height:82px;
    max-width:130px;
    max-height:82px;
    left:25px;
    top:0;
    object-fit:contain;
    object-position:center;
    display:block;
    z-index:2;
}

.cap{
    position:absolute;
    width:92px;
    height:82px;
    max-width:92px;
    max-height:82px;

    /* GESER CAP LEBIH KE KIRI */
    left:75px;

    top:1px;
    object-fit:contain;
    object-position:center;
    display:block;
    opacity:.90;
    z-index:3;
}

.nama-direktur{
    display:block;
    width:100%;
    font-size:13.5px;
    font-weight:bold;
    white-space:nowrap;
    margin:1px 0 0;
    line-height:1.2;
}

/* ================================
   FOOTER
================================ */
.footer{
    position:absolute;
    left:62px;
    right:62px;
    bottom:12px;
    border-top:1px solid #000;
    padding-top:3px;
    text-align:center;
    font-size:7.5px;
    line-height:1.25;
}

/* ================================
   DATABASE
================================ */
.database{margin-top:30px}

.database-toolbar{
    display:flex;
    gap:10px;
    align-items:center;
    margin-bottom:10px;
}

.database-toolbar select{
    max-width:220px;
}

.database table{
    width:100%;
    border-collapse:collapse;
}

.database th,
.database td{
    border:1px solid #ccc;
    padding:8px;
    font-size:13px;
    text-align:left;
}

.database th{
    background:#eee;
}

.empty-row{
    text-align:center !important;
    color:#777;
}

/* ================================
   RESPONSIVE PREVIEW
================================ */
@media(max-width:850px){
    body{padding:10px}
    .container{padding:15px}
    .preview-wrapper{padding:10px}
}

/* Print manual */
@media print{
    body{margin:0;padding:0;background:#fff}
    .container{display:none}
    .preview-wrapper{margin:0;padding:0;background:#fff}
    .preview{
        width:210mm;
        height:297mm;
        margin:0;
        padding:10mm 16mm 12mm;
    }
}
</style>
</head>

<body>
<div class="container">

<h2 class="page-title">Generator Surat Keterangan LSP LAS</h2>

<div class="form-group">
    <label>No Surat</label>
    <input id="nomor" type="text" placeholder="No. SUKET/001/LSP-LAS/IX/2026">
</div>

<div class="form-group">
    <label>Nama Asesi</label>
    <input id="nama" type="text" placeholder="Nama lengkap asesi">
</div>

<div class="form-group">
    <label>Skema</label>
    <select id="skema" onchange="toggleWelder()">
        <option>Fillet Welder</option>
        <option>Plate Welder</option>
        <option>Pipe Welder</option>
        <option>Group Leader</option>
        <option>Welding Inspector Basic</option>
        <option>Welding Foreman</option>
        <option>Welding Inspector Standard</option>
        <option>Welding Practitioner</option>
        <option>Welding Instructor</option>
        <option>Welding Specialist/Supervisor</option>
        <option>Welding Inspector Comprehensive</option>
        <option>Pipe Fitter</option>
        <option>Welding Engineer</option>
    </select>
</div>

<div class="form-group" id="welderBox">
    <label>Posisi / Proses</label>
    <input id="welder" list="welderList" placeholder="Contoh: 6G GTAW - SMAW">
    <datalist id="welderList"></datalist>
</div>

<div class="form-group">
    <label>Jenis TUK</label>
    <select id="jenisTuk">
        <option>Sewaktu</option>
        <option>Tempat Kerja</option>
        <option>Mandiri</option>
    </select>
</div>

<div class="form-group">
    <label>Nama TUK</label>
    <input id="tuk" list="tukList" type="text" placeholder="Contoh: BBPVP Bekasi">
    <datalist id="tukList"></datalist>
</div>

<div class="form-group">
    <label>Tanggal Uji</label>
    <input id="tanggal" type="date">
</div>

<div class="form-group">
    <label>Tanggal Tanda Tangan</label>
    <input id="ttd" type="date">
</div>

<button class="btn-generate" onclick="generate()">Generate</button>
<button class="btn-pdf" onclick="downloadPDF()">Download PDF</button>
<button class="btn-excel" onclick="exportExcel()">Export Excel</button>
<button class="btn-import" onclick="document.getElementById('importFile').click()">Import Excel</button>
<input type="file" id="importFile" accept=".xlsx,.xls" style="display:none" onchange="importExcel(event)">

<div class="preview-wrapper">
    <div class="preview" id="print"></div>
</div>

<div class="database">
    <h3>Database Surat</h3>

    <div class="database-toolbar">
        <label style="margin:0">Filter Bulan:</label>
        <select id="filterBulan" onchange="renderTable()">
            <option value="">Semua Bulan</option>
            <option value="1">Januari</option>
            <option value="2">Februari</option>
            <option value="3">Maret</option>
            <option value="4">April</option>
            <option value="5">Mei</option>
            <option value="6">Juni</option>
            <option value="7">Juli</option>
            <option value="8">Agustus</option>
            <option value="9">September</option>
            <option value="10">Oktober</option>
            <option value="11">November</option>
            <option value="12">Desember</option>
        </select>
    </div>

    <table>
        <thead>
            <tr>
                <th>No Surat</th>
                <th>Nama Asesi</th>
                <th>Skema</th>
                <th>Jenis TUK</th>
                <th>Nama TUK</th>
                <th>Tanggal Uji</th>
            </tr>
        </thead>
        <tbody id="tableBody"></tbody>
    </table>
</div>

</div>

<script>
/* =====================================================
   DATABASE
===================================================== */

let db = JSON.parse(localStorage.getItem("db_suket_lsp_las")) || [];
let welderData = JSON.parse(localStorage.getItem("welder_lsp_las")) || [];
let tukData = JSON.parse(localStorage.getItem("tuk_lsp_las")) || [];
let currentRecord = null;

/* =====================================================
   BULAN
===================================================== */

const namaBulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
];

const romawi = [
    "I","II","III","IV","V","VI",
    "VII","VIII","IX","X","XI","XII"
];

function bulanRomawi(bulan){
    return romawi[bulan - 1] || "";
}

function formatTanggal(tanggal){
    if(!tanggal) return "";

    const d = new Date(tanggal + "T00:00:00");

    return String(d.getDate()).padStart(2,"0")
        + " " + namaBulan[d.getMonth()]
        + " " + d.getFullYear();
}

function formatTanggalFile(tanggal){
    if(!tanggal) return "Tanggal";

    const d = new Date(tanggal + "T00:00:00");

    return String(d.getDate()).padStart(2,"0")
        + "-" + String(d.getMonth()+1).padStart(2,"0")
        + "-" + d.getFullYear();
}

/* =====================================================
   NOMOR SURAT
   Bulan + tahun selalu mengikuti kalender laptop/system
===================================================== */

function getNomorTerakhir(){
    let nomorTerakhir = Number(localStorage.getItem("last_sequence_suket_lsp_las")) || 0;

    db.forEach(item => {
        const match = String(item.nomor || "").match(/SUKET\/(\d+)/i);
        if(match){
            const angka = parseInt(match[1],10);
            if(!isNaN(angka) && angka > nomorTerakhir){
                nomorTerakhir = angka;
            }
        }
    });

    return nomorTerakhir;
}

function buatNomorBerikutnya(){
    const next = getNomorTerakhir() + 1;

    /* PENTING:
       Bulan dan tahun mengikuti kalender laptop,
       bukan tanggal Uji/Tanggal TTD.
    */
    const sekarang = new Date();
    const bulan = sekarang.getMonth() + 1;
    const tahun = sekarang.getFullYear();

    return "No. SUKET/"
        + String(next).padStart(3,"0")
        + "/LSP-LAS/"
        + bulanRomawi(bulan)
        + "/"
        + tahun;
}

function autoNomor(){
    document.getElementById("nomor").value = buatNomorBerikutnya();
}

/* =====================================================
   SKEMA / POSISI
===================================================== */

function toggleWelder(){
    const skema = document.getElementById("skema").value;
    const box = document.getElementById("welderBox");

    const perluPosisi =
        skema === "Fillet Welder" ||
        skema === "Plate Welder" ||
        skema === "Pipe Welder";

    box.style.display = perluPosisi ? "block" : "none";

    if(!perluPosisi){
        document.getElementById("welder").value = "";
    }
}

function saveWelder(value){
    value = String(value || "").trim();
    if(!value) return;

    if(!welderData.includes(value)){
        welderData.push(value);
        localStorage.setItem("welder_lsp_las", JSON.stringify(welderData));
    }
}

function loadWelder(){
    const list = document.getElementById("welderList");
    list.innerHTML = "";

    welderData.forEach(item => {
        const option = document.createElement("option");
        option.value = item;
        list.appendChild(option);
    });
}

/* =====================================================
   TUK
===================================================== */

function saveTuk(value){
    value = String(value || "").trim();
    if(!value) return;

    if(!tukData.includes(value)){
        tukData.push(value);
        localStorage.setItem("tuk_lsp_las", JSON.stringify(tukData));
    }
}

function loadTuk(){
    const list = document.getElementById("tukList");
    list.innerHTML = "";

    tukData.forEach(item => {
        const option = document.createElement("option");
        option.value = item;
        list.appendChild(option);
    });
}

/* =====================================================
   ESCAPE HTML
===================================================== */

function escapeHtml(value){
    return String(value ?? "")
        .replace(/&/g,"&amp;")
        .replace(/</g,"&lt;")
        .replace(/>/g,"&gt;")
        .replace(/"/g,"&quot;")
        .replace(/'/g,"&#039;");
}

/* =====================================================
   DATABASE
===================================================== */

function renderTable(){
    const tbody = document.getElementById("tableBody");
    const filter = document.getElementById("filterBulan").value;

    tbody.innerHTML = "";

    const filtered = db.filter(item => {
        if(!filter) return true;

        const tanggal = item.tanggal || "";
        if(!tanggal) return false;

        const d = new Date(tanggal + "T00:00:00");
        return String(d.getMonth()+1) === String(filter);
    });

    if(filtered.length === 0){
        tbody.innerHTML = `
            <tr>
                <td colspan="6" class="empty-row">Belum ada data surat.</td>
            </tr>
        `;
        return;
    }

    filtered.forEach(item => {
        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td>${escapeHtml(item.nomor || "")}</td>
            <td>${escapeHtml(item.nama || "")}</td>
            <td>${escapeHtml(item.skema || "")}</td>
            <td>${escapeHtml(item.jenisTuk || "")}</td>
            <td>${escapeHtml(item.tuk || "")}</td>
            <td>${escapeHtml(formatTanggal(item.tanggal || ""))}</td>
        `;

        tbody.appendChild(tr);
    });
}

/* =====================================================
   GENERATE
===================================================== */

function generate(){
    const nomor = document.getElementById("nomor").value.trim();
    const nama = document.getElementById("nama").value.trim();
    const skema = document.getElementById("skema").value;
    const welder = document.getElementById("welder").value.trim();
    const jenisTuk = document.getElementById("jenisTuk").value;
    const tuk = document.getElementById("tuk").value.trim();
    const tanggal = document.getElementById("tanggal").value;
    const ttd = document.getElementById("ttd").value;

    if(!nomor){
        alert("Nomor surat belum diisi.");
        return;
    }

    if(!nama){
        alert("Nama asesi belum diisi.");
        return;
    }

    if(!tanggal){
        alert("Tanggal Uji belum diisi.");
        return;
    }

    if(!ttd){
        alert("Tanggal Tanda Tangan belum diisi.");
        return;
    }

    if(
        (skema === "Fillet Welder" ||
         skema === "Plate Welder" ||
         skema === "Pipe Welder") &&
        !welder
    ){
        alert("Posisi / Proses belum diisi.");
        return;
    }

    if(!tuk){
        alert("Nama TUK belum diisi.");
        return;
    }

    saveWelder(welder);
    saveTuk(tuk);
    loadWelder();
    loadTuk();

    let skemaFinal = skema;
    if(welder){
        skemaFinal += " (" + escapeHtml(welder) + ")";
    }

    const nomorSafe = escapeHtml(nomor);
    const namaSafe = escapeHtml(nama);
    const jenisTukSafe = escapeHtml(jenisTuk);
    const tukSafe = escapeHtml(tuk);

    const tglUji = formatTanggal(tanggal);
    const tglTTD = formatTanggal(ttd);

    const html = `
        <div class="kop">
            <table>
                <tr>
                    <td class="logo-cell-left">
                        <img src="logo_lsp.png" class="logo-lsp" alt="Logo LSP" loading="eager" onerror="if(this.dataset.fallback!=='1'){this.dataset.fallback='1';this.src='logo_lsp.jpg';}else{this.style.display='none';}">
                    </td>

                    <td class="logo-cell-center">
                        <div class="nama-lembaga">
                            <h1>LEMBAGA SERTIFIKASI PROFESI LAS</h1>
                            <h2>(LSP-LAS)</h2>
                            <div class="izin">
                                Lisensi BNSP No. BNSP-LSP-024-ID
                                Tanggal 30 Januari 2026-2031,
                                Akta Notaris : Rismalena Kasri, S.H
                                No : 15 Tgl : 29 Mei 2007
                            </div>
                        </div>
                    </td>

                    <td class="logo-cell-right">
                        <img src="logo_bnsp.png" class="logo-bnsp" alt="Logo BNSP" loading="eager" onerror="if(this.dataset.fallback!=='1'){this.dataset.fallback='1';this.src='logo_bnsp.jpg';}else{this.style.display='none';}">
                    </td>
                </tr>
            </table>
        </div>

        <div class="judul">
            <h2>SURAT KETERANGAN</h2>
            <h3>${nomorSafe}</h3>
        </div>

        <div class="isi">
            Lembaga Sertifikasi Profesi Las (LSP-LAS)
            dengan ini menerangkan bahwa :

            <div class="nama-asesi">
                ${namaSafe}
            </div>

            Telah mengikuti Uji Kompetensi pada bidang Jasa Industri Pengelasan :

            <div class="skema">
                ${skemaFinal}
            </div>

            Yang diselenggarakan oleh Lembaga Sertifikasi
            Profesi Las (LSP-LAS) di TUK-LAS
            ${jenisTukSafe} ${tukSafe}
            pada tanggal ${escapeHtml(tglUji)}.

            <br><br>

            Dengan Hasil :

            <div class="hasil">
                Kompeten
            </div>

            Surat Keterangan ini dibuat sebagai pengganti sementara
            atas Sertifikat Kompetensi yang saat ini sedang dalam
            proses pengajuan ke BNSP.

            <br><br>

            Demikian Surat Keterangan ini dibuat untuk dipergunakan
            sebagaimana mestinya.

            <div class="ttd">
                <div class="ttd-area">
                    <div class="ttd-kota">
                        Jakarta, ${escapeHtml(tglTTD)}
                    </div>

                    <div class="ttd-jabatan">
                        Direktur
                    </div>

                    <div class="signature-visual">
                        <img src="sunoto.jpeg" class="ttd-img" alt="Tanda tangan" loading="eager" onerror="if(this.dataset.fallback!=='1'){this.dataset.fallback='1';this.src='sunoto.png';}else{this.style.display='none';}">
                        <img src="cap.jpeg" class="cap" alt="Cap" loading="eager" onerror="if(this.dataset.fallback!=='1'){this.dataset.fallback='1';this.src='cap.png';}else{this.style.display='none';}">
                    </div>

                    <div class="nama-direktur">
                        Ir. Sunoto Mudiantoro, M.T
                    </div>
                </div>
            </div>
        </div>

        <div class="footer">
            Casa Residence Jl. Bima II Blok A8 No.30A,
            Cijantung, Kec. Ps. Rebo,
            Kota Jakarta Timur,
            Daerah Khusus Ibukota Jakarta 13770.<br>
            Telp/WA 0819-17-100-200,
            Bank Mandiri KCP Jakarta Mampang Inspirasi
            Rek.No. 1177 00513 1970
        </div>
    `;

    document.getElementById("print").innerHTML = html;

    /* Snapshot surat yang benar untuk PDF */
    currentRecord = {
        nomor,
        nama,
        skema: skemaFinal.replace(/<[^>]*>/g,""),
        jenisTuk,
        tuk,
        tanggal,
        tanggalTTD: ttd,
        welder
    };

    /* Simpan database, tetapi jangan menggandakan nomor yang sama */
    const sudahAda = db.some(item => item.nomor === nomor);

    if(!sudahAda){
        db.push({
            nomor,
            nama,
            skema: skemaFinal.replace(/<[^>]*>/g,""),
            jenisTuk,
            tuk,
            tanggal,
            tanggalTTD: ttd,
            welder
        });
    }else{
        /* Kalau nomor sama, update datanya */
        const index = db.findIndex(item => item.nomor === nomor);
        if(index >= 0){
            db[index] = {
                ...db[index],
                nomor,
                nama,
                skema: skemaFinal.replace(/<[^>]*>/g,""),
                jenisTuk,
                tuk,
                tanggal,
                tanggalTTD: ttd,
                welder
            };
        }
    }

    /* Nomor manual yang dipakai menjadi acuan nomor berikutnya */
    const match = nomor.match(/SUKET\/(\d+)/i);
    if(match){
        const nomorSekarang = parseInt(match[1],10);
        const lastStored = Number(localStorage.getItem("last_sequence_suket_lsp_las")) || 0;

        if(!isNaN(nomorSekarang) && nomorSekarang > lastStored){
            localStorage.setItem(
                "last_sequence_suket_lsp_las",
                String(nomorSekarang)
            );
        }
    }

    localStorage.setItem(
        "db_suket_lsp_las",
        JSON.stringify(db)
    );

    renderTable();

    /* Setelah Generate, form pindah ke nomor berikutnya.
       Preview tetap menggunakan nomor surat yang baru saja dibuat. */
    autoNomor();

    alert("Surat berhasil di-generate.");
}

/* =====================================================
   PDF 1 HALAMAN A4
===================================================== */

async function waitImages(root){
    const images = [...root.querySelectorAll("img")];

    await Promise.all(
        images.map(img => {
            if(img.complete && img.naturalWidth > 0){
                return Promise.resolve();
            }

            return new Promise(resolve => {
                const done = () => {
                    img.removeEventListener("load", done);
                    img.removeEventListener("error", done);
                    resolve();
                };

                img.addEventListener("load", done, {once:true});
                img.addEventListener("error", done, {once:true});

                /* Jangan menunggu selamanya */
                setTimeout(done, 5000);
            });
        })
    );
}

async function downloadPDF(){
    const element = document.getElementById("print");

    if(!element || !element.innerHTML.trim()){
        alert("Silakan Generate surat terlebih dahulu.");
        return;
    }

    const record = currentRecord || {
        nomor: document.getElementById("nomor").value.trim(),
        tuk: document.getElementById("tuk").value.trim(),
        tanggal: document.getElementById("tanggal").value
    };

    const nomor = record.nomor || "No. SUKET/001/LSP-LAS";
    const tuk = record.tuk || "TUK";
    const tanggal = record.tanggal || "";

    const match = nomor.match(/SUKET\/(\d+)/i);
    const noOnly = match ? match[1].padStart(3,"0") : "001";

    const tukFormat = tuk
        .replace(/[\\/:*?"<>|]/g,"")
        .trim()
        .replace(/\s+/g,"-") || "TUK";

    const filename =
        `${noOnly}_SUKET_${tukFormat}_${formatTanggalFile(tanggal)}.pdf`;

    /* =====================================================
       CLONE A4 KHUSUS PDF
       Diposisikan di luar wrapper agar koordinat capture
       benar-benar mulai dari pojok A4.
    ===================================================== */
    const clone = element.cloneNode(true);

    clone.id = "pdfClone";
    clone.className = "preview pdf-clone";

    clone.style.position = "absolute";
    clone.style.left = "0";
    clone.style.top = "0";
    clone.style.margin = "0";
    clone.style.width = "794px";
    clone.style.height = "1123px";
    clone.style.padding = "30px 62px 78px";
    clone.style.background = "#fff";
    clone.style.overflow = "hidden";
    clone.style.zIndex = "-1";

    document.body.appendChild(clone);

    try{
        await waitImages(clone);

        /* Beri browser waktu untuk merender font + gambar */
        await new Promise(resolve => requestAnimationFrame(() => {
            requestAnimationFrame(resolve);
        }));

        const canvas = await html2canvas(clone,{
            scale:4,
            width:794,
            height:1123,
            windowWidth:794,
            windowHeight:1123,
            backgroundColor:"#ffffff",
            useCORS:true,
            allowTaint:false,
            logging:false,
            scrollX:0,
            scrollY:0,
            imageTimeout:15000
        });

        const imgData = canvas.toDataURL("image/jpeg",1.0);

        const {jsPDF} = window.jspdf;

        const pdf = new jsPDF({
            orientation:"portrait",
            unit:"mm",
            format:"a4",
            compress:true
        });

        /* Tepat satu halaman A4 */
        pdf.addImage(
            imgData,
            "JPEG",
            0,
            0,
            210,
            297,
            undefined,
            "FAST"
        );

        pdf.save(filename);

    }catch(error){
        console.error("PDF ERROR:", error);
        alert(
            "PDF gagal dibuat. Pastikan file logo_lsp.png, logo_bnsp.png, sunoto.jpeg/sunoto.png dan cap.jpeg/cap.png ada di folder yang sama."
        );
    }finally{
        clone.remove();
    }
}

/* =====================================================
   EXPORT EXCEL
===================================================== */

function exportExcel(){
    if(db.length === 0){
        alert("Database masih kosong.");
        return;
    }

    const data = db.map(item => ({
        "No Surat": item.nomor || "",
        "Nama Asesi": item.nama || "",
        "Skema": item.skema || "",
        "Jenis TUK": item.jenisTuk || "",
        "Nama TUK": item.tuk || "",
        "Tanggal Uji": item.tanggal || "",
        "Tanggal TTD": item.tanggalTTD || "",
        "Posisi / Proses": item.welder || ""
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb,ws,"Database Surat");
    XLSX.writeFile(wb,"Database_Suket_LSP_LAS.xlsx");
}

/* =====================================================
   IMPORT EXCEL
===================================================== */

function importExcel(event){
    const file = event.target.files[0];
    if(!file) return;

    const reader = new FileReader();

    reader.onload = function(e){
        try{
            const data = new Uint8Array(e.target.result);

            const workbook = XLSX.read(data,{type:"array"});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            const rows = XLSX.utils.sheet_to_json(sheet,{defval:""});

            if(!rows.length){
                alert("Excel tidak mempunyai data.");
                return;
            }

            const imported = rows.map(row => ({
                nomor:
                    row["No Surat"] ||
                    row["Nomor"] ||
                    row["No. Surat"] ||
                    "",

                nama:
                    row["Nama Asesi"] ||
                    row["Nama"] ||
                    "",

                skema:
                    row["Skema"] ||
                    "",

                jenisTuk:
                    row["Jenis TUK"] ||
                    "",

                tuk:
                    row["Nama TUK"] ||
                    row["TUK"] ||
                    "",

                tanggal:
                    row["Tanggal Uji"] ||
                    row["Tanggal"] ||
                    "",

                tanggalTTD:
                    row["Tanggal TTD"] ||
                    "",

                welder:
                    row["Posisi / Proses"] ||
                    ""
            })).filter(item => item.nomor);

            imported.forEach(item => {
                const index = db.findIndex(x => x.nomor === item.nomor);

                if(index >= 0){
                    db[index] = {...db[index],...item};
                }else{
                    db.push(item);
                }

                const match = String(item.nomor).match(/SUKET\/(\d+)/i);
                if(match){
                    const angka = parseInt(match[1],10);
                    const last = Number(localStorage.getItem("last_sequence_suket_lsp_las")) || 0;

                    if(!isNaN(angka) && angka > last){
                        localStorage.setItem(
                            "last_sequence_suket_lsp_las",
                            String(angka)
                        );
                    }
                }

                if(item.welder) saveWelder(item.welder);
                if(item.tuk) saveTuk(item.tuk);
            });

            localStorage.setItem(
                "db_suket_lsp_las",
                JSON.stringify(db)
            );

            loadWelder();
            loadTuk();
            renderTable();
            autoNomor();

            alert("Import Excel berhasil.");

        }catch(error){
            console.error(error);
            alert("Gagal membaca file Excel.");
        }finally{
            event.target.value = "";
        }
    };

    reader.readAsArrayBuffer(file);
}

/* =====================================================
   TANGGAL TTD DEFAULT = HARI INI
   Tetap bisa diubah manual oleh pengguna.
===================================================== */

function setTanggalTTDHariIni(){
    const input = document.getElementById("ttd");

    if(!input.value){
        const sekarang = new Date();
        const tahun = sekarang.getFullYear();
        const bulan = String(sekarang.getMonth() + 1).padStart(2,"0");
        const hari = String(sekarang.getDate()).padStart(2,"0");

        input.value = `${tahun}-${bulan}-${hari}`;
    }
}

/* =====================================================
   INIT
===================================================== */

setTanggalTTDHariIni();
loadWelder();
loadTuk();
toggleWelder();
renderTable();
autoNomor();
</script>
</body>
</html>
