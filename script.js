/* =========================================================
   GudangKita - Sistem Manajemen Stok & Peminjaman Gudang
   Data disimpan di localStorage (berfungsi sebagai "database"
   sisi klien). Semua modul: Barang, Peminjaman/Pengembalian,
   Barang Rusak/MT, Pembelian, Pengajuan Pembelian, Barang
   Hilang, dan Manajemen User.
   ========================================================= */

const DB_KEY = 'gudangkita_db_v1';
const SESSION_KEY = 'gudangkita_session_v1';

/* ---------- Util umum ---------- */
function uid(prefix) {
    return prefix + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}
function todayStr() {
    return new Date().toISOString().slice(0, 10);
}
function fmtDate(d) {
    if (!d) return '-';
    const dt = new Date(d);
    if (isNaN(dt)) return d;
    return dt.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
}
function fmtDateTime(d) {
    const dt = new Date(d);
    if (isNaN(dt)) return d;
    return dt.toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}
function fmtRupiah(n) {
    n = Number(n) || 0;
    return 'Rp ' + n.toLocaleString('id-ID');
}
function escapeHtml(str) {
    if (str === undefined || str === null) return '';
    return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ---------- Database (localStorage) ----------
   ITEMS_DATA_VERSION dinaikkan setiap kali daftar barang bawaan (import
   dari file Excel/data asli) diperbarui, supaya browser yang sudah
   punya data lama otomatis ikut diperbarui tanpa kehilangan riwayat
   peminjaman/pembelian/dll yang sudah dibuat. */
const ITEMS_DATA_VERSION = 2;
const IMPORTED_ITEMS = [
            { id: 'AK0001', name: 'AIR STAPPLER U', category: 'Alat Kerja', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0002', name: 'BRAD NAILER I', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0003', name: 'DONGCHENG BATRE', category: 'Alat Kerja', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0004', name: 'CHARGER DONGCHENG', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0005', name: 'BATRE DONGCHENG', category: 'Alat Kerja', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0006', name: 'BOR BATRE', category: 'Alat Kerja', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0007', name: 'CHARGER BOR BATRE', category: 'Alat Kerja', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0008', name: 'BATRE BOR', category: 'Alat Kerja', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0009', name: 'MESIN COMPRESSOR', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0010', name: 'HOTGUN', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0011', name: 'JIGSAW', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0012', name: 'MESIN AMPLAS', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0013', name: 'BOR LISTRIK', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0014', name: 'MESIN LAS', category: 'Alat Kerja', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0015', name: 'ROUTER', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0016', name: 'MESIN POTONG MEJA 7\'', category: 'Alat Kerja', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0017', name: 'MESIN POTONG DUDUK 7\'', category: 'Alat Kerja', stock: 0, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'AK0018', name: 'MESIN POTONG MEJA 14\'', category: 'Alat Kerja', stock: 0, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00001', name: 'GERABAH MILL&BAY 60CM PUTIH COKELAT', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00002', name: 'GERABAH MILL&BAY 60CM PUTIH', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00003', name: 'GERABAH MILL&BAY 30CM MERAH BAMBU UKIR', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00004', name: 'GERABAH MILL&BAY 35CM MOTIF UKIR', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00005', name: 'GERABAH MILL&BAY 50CM PUTIH BOTOL', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00006', name: 'GERABAH MILL&BAY 40CM GRADASI MERAH KUNING', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00007', name: 'GERABAH MILL&BAY 40CM GRADASI HIJAU KUNING', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00008', name: 'GERABAH MILL&BAY 120CM COKELAT', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00009', name: 'PETI KAYU 31X23X15', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00010', name: 'PETI KAYU 50X38X30', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00011', name: 'PETI KAYU 46X32X25', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00012', name: 'PETI KAYU 60X40X35', category: 'Dekorasi', stock: 7, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00013', name: 'LILIN 20CM ORIGINAL', category: 'Dekorasi', stock: 18, unit: 'pcs', location: 'Gudang Utama', rack: '1.18', minStock: 1 },
            { id: 'D00014', name: 'LILIN 28CM CREAM', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00015', name: 'LILIN 30CM RED', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00016', name: 'LILIN 28CM BLACK', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00017', name: 'LILIN 15CM BLACK', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00018', name: 'LILIN 10CM GREY', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00019', name: 'LILIN 20CM ROSE', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00020', name: 'LILIN 20CM ORANGE', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00021', name: 'LILIN 20CM APPLE', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00022', name: 'LILIN 20CM GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00023', name: 'LILIN 4CM RED', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00024', name: 'LILIN 20CM RED', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00025', name: 'LILIN BATRE WHITE', category: 'Dekorasi', stock: 53, unit: 'pcs', location: 'Gudang Utama', rack: '1,7,18', minStock: 1 },
            { id: 'D00026', name: 'LILIN BATRE RED', category: 'Dekorasi', stock: 22, unit: 'pcs', location: 'Gudang Utama', rack: '1', minStock: 1 },
            { id: 'D00027', name: 'FRAME HOLDER 10X15 2FACE WHITE', category: 'Dekorasi', stock: 7, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00028', name: 'FRAME HOLDER 10X15 2FACE BLACK', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00029', name: 'PHOTO FRAME 4R ROSE SILVER', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00030', name: 'PHOTO FRAME 4R ROSE GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00031', name: 'PHOTO FRAME 4R GREEN MATCHA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00032', name: 'PHOTO FRAME 4R SILVER', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00033', name: 'PHOTO FRAME 3R UKIR GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00034', name: 'PHOTO FRAME 10X10 KAIN UNGU', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00035', name: 'MINIATUR TOPENG 25CM SILVER', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00036', name: 'PHOTOBOOTH PROPS', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '2.26', minStock: 1 },
            { id: 'D00037', name: 'CUPCAKE STAND 3 TIERS DELICIA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00038', name: 'POT PAGAR DEKORASI 20X10X10 WHITE', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '2', minStock: 1 },
            { id: 'D00039', name: 'TOPLES KACA 20CM', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00040', name: 'ONE TWO CUPS 20CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00041', name: 'ONE TWO CUPS 15CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00042', name: 'ONE TWO CUPS 12CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00043', name: 'FISHBOWL DIA20CM T20CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00044', name: 'VAS KERAMIK 0104S', category: 'Dekorasi', stock: 15, unit: 'pcs', location: 'Gudang Utama', rack: '3,16,28', minStock: 1 },
            { id: 'D00045', name: 'CANDLE CASE', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00046', name: 'KACA SELONGSONG DIA8CM T9CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00047', name: 'VAS KACA BOWL DIA8CM T8CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00048', name: 'GELAS CARAFE T19CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00049', name: 'GELAS CARAFE T20CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3.16', minStock: 1 },
            { id: 'D00050', name: 'VAS VILJESTARK T17CM IKEA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '3', minStock: 1 },
            { id: 'D00051', name: 'MAINAN SPIRAL RAINBOW', category: 'Dekorasi', stock: 58, unit: 'pcs', location: 'Gudang Utama', rack: '4', minStock: 1 },
            { id: 'D00052', name: 'RUBIK', category: 'Dekorasi', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '4', minStock: 1 },
            { id: 'D00053', name: 'PIRINGAN HITAM', category: 'Dekorasi', stock: 39, unit: 'pcs', location: 'Gudang Utama', rack: '5.36', minStock: 1 },
            { id: 'D00054', name: 'KASET', category: 'Dekorasi', stock: 37, unit: 'pcs', location: 'Gudang Utama', rack: '5', minStock: 1 },
            { id: 'D00055', name: 'VAS KERAMIK 011S', category: 'Dekorasi', stock: 45, unit: 'pcs', location: 'Gudang Utama', rack: '6,15,16,28', minStock: 1 },
            { id: 'D00056', name: 'TUSCANY GLASS', category: 'Dekorasi', stock: 17, unit: 'pcs', location: 'Gudang Utama', rack: '6,29,44', minStock: 1 },
            { id: 'D00057', name: 'BRANDY GLASS DIA13CM T18', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '6', minStock: 1 },
            { id: 'D00058', name: 'LENTERA PUTIH T20CM IKEA', category: 'Dekorasi', stock: 23, unit: 'pcs', location: 'Gudang Utama', rack: '7,10,14', minStock: 1 },
            { id: 'D00059', name: 'ARCOROC BOWLS BLACK DIA8CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00060', name: 'ARCOROC BOWLS BLACK DIA14CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00061', name: 'ARCOROC BOWLS BLACK DIA23CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00062', name: 'ARCOROC PLATE BLACK', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00063', name: 'BOTOL KACA COKELAT T19CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '8.22', minStock: 1 },
            { id: 'D00064', name: 'BOTOL KACA GEPENG COKELAT T19CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00065', name: 'CANGKIR VICENZA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00066', name: 'COLORADO GLASS', category: 'Dekorasi', stock: 30, unit: 'pcs', location: 'Gudang Utama', rack: '8,tba', minStock: 1 },
            { id: 'D00067', name: 'DRINKING JAR', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00068', name: 'GELAS ACAR', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '8.22', minStock: 1 },
            { id: 'D00069', name: 'CLEAR JAR', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00070', name: 'PURPLE SPECKLED GLASS', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00071', name: 'CHANDELIER CRYSTAL DIA6CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00072', name: 'BLUE POTTERY PLATE', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '8', minStock: 1 },
            { id: 'D00073', name: 'MUTIARA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '9', minStock: 1 },
            { id: 'D00074', name: 'KAIN BIRU', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: '11,SK2', minStock: 1 },
            { id: 'D00075', name: 'KAIN KERLAP KERLIP BIRU', category: 'Dekorasi', stock: 9, unit: 'pcs', location: 'Gudang Utama', rack: '11,SK2', minStock: 1 },
            { id: 'D00076', name: 'KAIN KERLAP KERLIP GOLD', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '12', minStock: 1 },
            { id: 'D00077', name: 'CANDLE HOLDER T45CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '13', minStock: 1 },
            { id: 'D00078', name: 'LAMPION', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00079', name: 'PAPERCUP', category: 'Dekorasi', stock: 60, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00080', name: 'KOTAK SESERAHAN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00081', name: 'STICK ESKRIM PLASTIK', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00082', name: 'LILIN ULANG TAHUN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00083', name: 'KERTAS KADO', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00084', name: 'RANSEL ANAK', category: 'Dekorasi', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00085', name: 'PITA SERUT', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '15', minStock: 1 },
            { id: 'D00086', name: 'GELAS CARAFE T15CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '16', minStock: 1 },
            { id: 'D00087', name: 'BORRBY LENTERA 28CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '16', minStock: 1 },
            { id: 'D00088', name: 'GOTTGORA LENTERA 25CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '16', minStock: 1 },
            { id: 'D00089', name: 'SANGKAR DEKORASI 15CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '16', minStock: 1 },
            { id: 'D00090', name: 'SANGKAR DEKORASI 20CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '16', minStock: 1 },
            { id: 'D00091', name: 'SANGKAR DEKORASI 15CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '16', minStock: 1 },
            { id: 'D00092', name: 'ALOHA TOPLES 4L', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '17', minStock: 1 },
            { id: 'D00093', name: 'DREAM CATCHER WHITE', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '17', minStock: 1 },
            { id: 'D00094', name: 'SERVER ROUND 2LAYER INFORMA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00095', name: 'MC KATY PERANGKAT MAKAN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00096', name: 'GLASS CAKE DIA30CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00097', name: 'VINTER CANDLE STICK 20CM IKEA', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00098', name: 'LAMPU TIDUR HIAS 20CM', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00099', name: 'GLASS CANDLE HOLDER 15CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00100', name: 'GLASS CANDLE HOLDER 10CM', category: 'Dekorasi', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00101', name: 'GLASS CANDLE HOLDER SQUARE 5CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00102', name: 'CANDLE DISH DIA18CM IKEA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00103', name: 'LILIN POETABLE BATRE T6CM MRDIY', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00104', name: 'LILIN 20CM PINK', category: 'Dekorasi', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '18', minStock: 1 },
            { id: 'D00105', name: 'PITA JEPANG HIJAU', category: 'Dekorasi', stock: 42, unit: 'pcs', location: 'Gudang Utama', rack: '19', minStock: 1 },
            { id: 'D00106', name: 'PITA JEPANG BIRU', category: 'Dekorasi', stock: 41, unit: 'pcs', location: 'Gudang Utama', rack: '19', minStock: 1 },
            { id: 'D00107', name: 'PITA JEPANG PINK', category: 'Dekorasi', stock: 40, unit: 'pcs', location: 'Gudang Utama', rack: '19', minStock: 1 },
            { id: 'D00108', name: 'CAKE LAYER 3 PUTIH', category: 'Dekorasi', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00109', name: 'CAKE LAYER 3 HIJAU', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00110', name: 'BOLA DISKO BESAR', category: 'Dekorasi', stock: 15, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00111', name: 'BOLA DISKO KECIL', category: 'Dekorasi', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00112', name: 'TELENAN BATANG POHON', category: 'Dekorasi', stock: 9, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00113', name: 'SALT PAPER CASE BENTUK BURUNG', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00114', name: 'PATUNG RUSA KUNO', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '20', minStock: 1 },
            { id: 'D00115', name: 'DECORATIVE VASE INFORMA T50CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '21', minStock: 1 },
            { id: 'D00116', name: 'LABU UKUR 1000ML', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00117', name: 'BEAKER GLASS 500ML', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00118', name: 'MEASURING CYLINDER', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00119', name: 'ERLENMEYER 100ML', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00120', name: 'ERLENMEYER 250ML', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00121', name: 'KORKEN BOTOL T20CM IKEA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00122', name: 'CYCLONE GLASS T21CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00123', name: 'KIG BOWL BLUE DIA12CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00124', name: 'KIG BOWL BLUE DIA7CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00125', name: 'TABLE NUMBER 10X20', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00126', name: 'TABLE NUMBER 15X10', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00127', name: 'TABLE NUMBER 15X20', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00128', name: 'STAND CARD 10X20', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00129', name: 'PIPET', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '22', minStock: 1 },
            { id: 'D00130', name: 'KAIN HITAM EVENT', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '23', minStock: 1 },
            { id: 'D00131', name: 'STETOSKOP HIJAU', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00132', name: 'FEATHER PEN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00133', name: 'INTRODUCTORY ORGANIC SET', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00134', name: 'BOTOL SCOT DURAN 400ML', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00135', name: 'BOTOL SCOT DURAN 200ML', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00136', name: 'BOTOL ASI', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00137', name: 'SUNTIKAN MAINAN', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '24', minStock: 1 },
            { id: 'D00138', name: 'PASIR PANTAI KERANG', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '25', minStock: 1 },
            { id: 'D00139', name: 'CHALKBOARD 20X15', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '26', minStock: 1 },
            { id: 'D00140', name: 'CHALKBOARD 30X20', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '26', minStock: 1 },
            { id: 'D00141', name: 'LAMPU TIRAI KETUPAT', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '27', minStock: 1 },
            { id: 'D00142', name: 'GODAFTON IKEA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00143', name: 'CANDLE FLAMELESS TWILIGHT', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00144', name: 'VASE GLASS ACE T25CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00145', name: 'STOPEN IKEA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00146', name: 'LJUSANDE IKEA', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00147', name: 'CONE TREE GOLD INFORMA T23CM', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '29.36', minStock: 1 },
            { id: 'D00148', name: 'WHISKEY DECANTER', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00149', name: 'GLASS CANDLE HOLDER T15CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00150', name: 'VTG COBALT TEAPOT BLUE', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00151', name: 'VTG COBALT MUG BLUE', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00152', name: 'VTG COBALT CREAMER CASE BLUE', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00153', name: 'VTG COBALT GRAVY BOAT BLUE', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00154', name: 'VTG COBALT MUG PLATE BLUE', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00155', name: 'UMBUL UMBUL BALI T50CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00156', name: 'PHOTO FRAME 5R TOSCA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00157', name: 'FISHBOWL T20CM DIA9CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00158', name: 'VINTAGE ADDISON RUSS LONDON CLOCK', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00159', name: 'GANTUNGAN BOHLAM NATAL', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00160', name: 'GANTUNGAN POHON NATAL HIJAU', category: 'Dekorasi', stock: 72, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00161', name: 'GANTUNGAN KUDA SILVER', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00162', name: 'SLING MERAH 4M', category: 'Dekorasi', stock: 7, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00163', name: 'SLING MERAH 2M', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00164', name: 'JARING WHITE', category: 'Dekorasi', stock: 7, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00165', name: 'JARING GOLD', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00166', name: 'TANGKAI CHERRY', category: 'Dekorasi', stock: 24, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00167', name: 'TANGKAI DAUN GREEN', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00168', name: 'TANGKAI DAUN GOLD', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00169', name: 'TANGKAI DAUN CREAM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '30', minStock: 1 },
            { id: 'D00170', name: 'POHON NATAL MINI 10CM', category: 'Dekorasi', stock: 11, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00171', name: 'POHON NATAL MINI 15CM', category: 'Dekorasi', stock: 9, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00172', name: 'POHON NATAL MINI 20CM', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00173', name: 'POHON NATALMINI 30CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00174', name: 'GANTUNGAN PINTU NATAL DIA30CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00175', name: 'GANTUNGAN PINTU NATAL DIA60CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00176', name: 'PITA MAROON BLUDURU 30CM', category: 'Dekorasi', stock: 44, unit: 'pcs', location: 'Gudang Utama', rack: '31', minStock: 1 },
            { id: 'D00177', name: 'GANTUNGAN KAOSKAKI BESAR 40CM', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00178', name: 'GANTUNGAN KAOSAKAKI SEDANG 30CM', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00179', name: 'GANTUNGAN KAOSKAKI KECIL 25CM', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00180', name: 'GANTUNGAN POHON NATAL SILVER', category: 'Dekorasi', stock: 23, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00181', name: 'GANTUNGAN MERRY CHRISMAST 60X20', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00182', name: 'GANTUNGAN MERRY CHRISMAST 40X15', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00183', name: 'GANTUNGAN MERRY CHRISMAST 30X15', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00184', name: 'GANTUNGAN MERRY CHRISMAST 15X7', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00185', name: 'GARLAND ARTIFICIAL SILVER 120CM', category: 'Dekorasi', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00186', name: 'GARLAND ARTIFICIAL GOLD 120CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00187', name: 'GANTUNGAN PERMEN PAYUNG MERAH PUTIH', category: 'Dekorasi', stock: 113, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00188', name: 'GANTUNGAN PERMEN BALON MERAH PUTIH', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00189', name: 'GANTUNGAN PERMEN PAYUNG MERAH', category: 'Dekorasi', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00190', name: 'GANTUNGAN PERMEN PAYUNG GOLD', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00191', name: 'GANTUNGAN GINGERMAN MERAH', category: 'Dekorasi', stock: 9, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00192', name: 'BOLA CHERRY', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '32', minStock: 1 },
            { id: 'D00193', name: 'NUTRACKER BESAR', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '33', minStock: 1 },
            { id: 'D00194', name: 'NUTRACKER SEDANG', category: 'Dekorasi', stock: 11, unit: 'pcs', location: 'Gudang Utama', rack: '33', minStock: 1 },
            { id: 'D00195', name: 'NUTRACKER KECIL', category: 'Dekorasi', stock: 11, unit: 'pcs', location: 'Gudang Utama', rack: '33', minStock: 1 },
            { id: 'D00196', name: 'GANTUNGAN NUTRACKER SEDANG', category: 'Dekorasi', stock: 19, unit: 'pcs', location: 'Gudang Utama', rack: '33', minStock: 1 },
            { id: 'D00197', name: 'GANTUNGAN NUTRACKER KECIL', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '33', minStock: 1 },
            { id: 'D00198', name: 'LAMPU THUMBLER WARM (KABEL PUTIH)', category: 'Dekorasi', stock: 46, unit: 'pcs', location: 'Gudang Utama', rack: '34', minStock: 1 },
            { id: 'D00199', name: 'MILL&BAY WATER HYACINTH T30 DIA30', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00200', name: 'MILL&BAY WATER HYACINTH T25 DIA30', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00201', name: 'MILL&BAY WATER HYACINTH T35 DIA15', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00202', name: 'MILL&BAY WATER HYACINTH T35 DIA30', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00203', name: 'POT BUNGA ROTAN T45', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00204', name: 'LAMPU GANTUNG KRISTAL', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00205', name: 'KAIN KERLAP KERLIP SILVER', category: 'Dekorasi', stock: 21, unit: 'pcs', location: 'Gudang Utama', rack: '35,SK5', minStock: 1 },
            { id: 'D00206', name: 'FIGURINE INFORMA', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00207', name: 'UNIQUE BOTTLE', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00208', name: 'ONE TWO CUPS PLATE GOLD DIA25', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00209', name: 'ONYX DINNER PLATE', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00210', name: 'CANDLE HOLDER KUNINGAN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00211', name: 'PATAKA DESK T35', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00212', name: 'RESERVED TABLE SIGN 15X5', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '36', minStock: 1 },
            { id: 'D00213', name: 'TAMBANG DEKORASI BAJAK LAUT', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00214', name: 'JARING IKAN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00215', name: 'TEROPONG BAJAK LAUT', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00216', name: 'PEDANG BAJAK LAUT', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00217', name: 'TUTUP MATA BAJAK LAUT', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00218', name: 'BAN PELAMPUNG DUMMY', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00219', name: 'SETIR KAPAL LAUT KECIL', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00220', name: 'KERANG', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '37', minStock: 1 },
            { id: 'D00221', name: 'HANGING SKELETON', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00222', name: 'POHON AGLONEMA 90CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00223', name: 'BUNGA POT GOLD 65CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00224', name: 'TEDDY BEAR 150CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00225', name: 'TEDDY BEAR 90CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00226', name: 'TEDDY BEAR 60CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00227', name: 'STANDING TEDDY BEAR 65CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00228', name: 'STANDING MERRY MEYER 55CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00229', name: 'SANTA CLAUSE WHITE GOLD 30CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00230', name: 'TABLE LANTERN MAROCO INFORMA 50CM', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00231', name: 'TABLE LANTERN MAROCO INFORMA 40CM', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00232', name: 'TABLE LANTERN MAROCO INFORMA 35CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00233', name: 'TABLE LANTERN MAROCO INFORMA 30CM', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00234', name: 'STEROFOAM RADIO', category: 'Dekorasi', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00235', name: 'STEROFOAM KAKTUS', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00236', name: 'KOTAK KADO NATAL', category: 'Dekorasi', stock: 50, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00237', name: 'KOTAK KADO NATAL BESEK', category: 'Dekorasi', stock: 25, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00238', name: 'KOTAK KADO NATAL TRANSPARAN', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00239', name: 'KOTAK KADO NATAL KERTAS MINYAK', category: 'Dekorasi', stock: 28, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00240', name: 'GLOBE DIA30CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00241', name: 'GLOBE DIA20CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00242', name: 'POT GERABAH COKELAT 35CM', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00243', name: 'SAYUR BUAH ARTIFICIAL', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00244', name: 'KOTAK POLYFOAM 20X20', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00245', name: 'KOTAK POLYFOAM 40X40', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00246', name: 'POT KECIL DAUN RUMBAI', category: 'Dekorasi', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00247', name: 'HMR BULAN SABIT 40CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00248', name: 'HMR BULAN SABIT 30CM', category: 'Dekorasi', stock: 18, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00249', name: 'HMR LAMPION 50CM', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00250', name: 'HMR LAMPION 40CM', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00251', name: 'GLITTERY HIASAN NATAL', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00252', name: 'GLITTERY HIASAN RAMADHAN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00253', name: 'TOTEM JANGKAR 60CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00254', name: 'PAGAR PLASTIK 75X45', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00255', name: 'PAGAR PLASTIK 60X30', category: 'Dekorasi', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00256', name: 'DRIED PHALANS', category: 'Dekorasi', stock: 22, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00257', name: 'ANYAMAN KECIL NATAL MIX', category: 'Dekorasi', stock: 22, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00258', name: 'PALEM KERING COKELAT', category: 'Dekorasi', stock: 11, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00259', name: 'PALEM KERING PINK', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00260', name: 'OBOR BAMBU 90CM', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00261', name: 'HMR KOTAK IMLEK 30X30', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00262', name: 'LAMPION IMLEK PUTIH', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00263', name: 'LAMPION IMLEK MERAH', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00264', name: 'LAMPION IMLEK KUNING', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00265', name: 'LAMPION IMLEK BIRU', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00266', name: 'BUNGA MAWAR KERTAS MIX', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00267', name: 'XMAS DEER LIGHT STANDING 70X30', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00268', name: 'ANYAMAN KETUPAT BESAR', category: 'Dekorasi', stock: 39, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00269', name: 'ANYAMAN KETUPAT SEDANG', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00270', name: 'ANYAMAN KETUPAT KECIL', category: 'Dekorasi', stock: 80, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00271', name: 'LAMPU THUMBLER WHITE (KABEL HITAM)', category: 'Dekorasi', stock: 35, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00272', name: 'LAMPU THUMBLER WARM (KABEL HITAM)', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00273', name: 'FRAME PUTIH 45X34', category: 'Dekorasi', stock: 24, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00274', name: 'MATRAS PUZZLE 30X30', category: 'Dekorasi', stock: 90, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00275', name: 'XMAS POLY VILLAGE TRAIN WITH MUSIC', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00276', name: 'XMAS POLY VILLAGE FERRIS WHEEL', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00277', name: 'LAMPU TANGKAI HIAS BUNGA PUTIH', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00278', name: 'GOTTGORA LANTERN 40CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00279', name: 'BULU ANGSA CREAM 50CM', category: 'Dekorasi', stock: 24, unit: 'pcs', location: 'Gudang Utama', rack: '38', minStock: 1 },
            { id: 'D00280', name: 'BULU ANGSA PUTIH 50CM', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '38', minStock: 1 },
            { id: 'D00281', name: 'BULU ANGSA COKELAT 50CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '38', minStock: 1 },
            { id: 'D00282', name: 'BULU ANGSA HITAM 50CM', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '38', minStock: 1 },
            { id: 'D00283', name: 'BULU ANGSA PUTIH 30CM', category: 'Dekorasi', stock: 11, unit: 'pcs', location: 'Gudang Utama', rack: '38', minStock: 1 },
            { id: 'D00284', name: 'LAMPU GANTUNG 10L WARM', category: 'Dekorasi', stock: 40, unit: 'pcs', location: 'Gudang Utama', rack: '39', minStock: 1 },
            { id: 'D00285', name: 'ROTATING GENIE LAMP', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '39', minStock: 1 },
            { id: 'D00286', name: 'VINTAGE GLASS BOTTLE 15CM', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '40 (MBAK MEN)', minStock: 1 },
            { id: 'D00287', name: 'NORDIC BOTTLE 13CM', category: 'Dekorasi', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '40 (MBAK MEN)', minStock: 1 },
            { id: 'D00288', name: 'BOLA NATAL MERAH MIX', category: 'Dekorasi', stock: 404, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00289', name: 'BOLA NATAL GOLD MIX', category: 'Dekorasi', stock: 293, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00290', name: 'BOLA NATAL SILVER MIX', category: 'Dekorasi', stock: 323, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00291', name: 'BOLA NATAL HIJAU MIX', category: 'Dekorasi', stock: 55, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00292', name: 'BOLA NATAL BIRU MIX', category: 'Dekorasi', stock: 160, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00293', name: 'TAMPAH DIA60', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00294', name: 'TAMPAH DIA55', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00295', name: 'TAMPAH DIA40', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00296', name: 'PORSELEN CHINA', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00297', name: 'LED GLASS CANDLE 15CM', category: 'Dekorasi', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '41', minStock: 1 },
            { id: 'D00298', name: 'BANTAL MOTIF MIX', category: 'Dekorasi', stock: 35, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00299', name: 'KARPET ROTAN', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00300', name: 'TENDA INDIAN', category: 'Dekorasi', stock: 7, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00301', name: 'POHON NATAL PNP', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00302', name: 'LAMPU HIAS XMAS RED', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00303', name: 'LAMPU GANTUNG HIAS', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00304', name: 'LAMPU TAMAN', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00305', name: 'POT BUNGA ROTAN T38', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00306', name: 'KERANJANG PIKNIK ROTAN', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00307', name: 'KERANJANG PIKNIK ROTAN DIA30', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00308', name: 'STANDING POHON NATAL 90CM', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00309', name: 'SELIMUT SALUR PUTIH 180X200', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK1', minStock: 1 },
            { id: 'D00310', name: 'SELIMUT MOTIF CATUR 180X40', category: 'Dekorasi', stock: 15, unit: 'pcs', location: 'Gudang Utama', rack: 'SK1', minStock: 1 },
            { id: 'D00311', name: 'KAIN MOTIF AWAN EX PELINDO', category: 'Dekorasi', stock: 17, unit: 'pcs', location: 'Gudang Utama', rack: 'SK1', minStock: 1 },
            { id: 'D00312', name: 'KAIN MOTIF BUNGA TROPIC', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: 'SK1', minStock: 1 },
            { id: 'D00313', name: 'TAPLAK MOTIF AWAN EX PELINDO', category: 'Dekorasi', stock: 22, unit: 'pcs', location: 'Gudang Utama', rack: 'SK1', minStock: 1 },
            { id: 'D00314', name: 'KARPET MOTIF SONGKET', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK2', minStock: 1 },
            { id: 'D00315', name: 'KAIN MOTIF SONGKET', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'SK2', minStock: 1 },
            { id: 'D00316', name: 'KARPET MOTIF PAKISTAN DUPATTA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK2', minStock: 1 },
            { id: 'D00317', name: 'TAPLAK GREY 140X40', category: 'Dekorasi', stock: 26, unit: 'pcs', location: 'Gudang Utama', rack: 'SK2', minStock: 1 },
            { id: 'D00318', name: 'TAPLAK MOTIF KOTAK KUNING PUTIH 200X50', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: 'SK2', minStock: 1 },
            { id: 'D00319', name: 'KAIN PUTIH BESAR', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK3', minStock: 1 },
            { id: 'D00320', name: 'KAIN PUTIH RENDA', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: 'SK3', minStock: 1 },
            { id: 'D00321', name: 'KAIN PUTIH POTONG', category: 'Dekorasi', stock: 32, unit: 'pcs', location: 'Gudang Utama', rack: 'SK3', minStock: 1 },
            { id: 'D00322', name: 'KAIN PUTIH', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: 'SK3,SK8', minStock: 1 },
            { id: 'D00323', name: 'KAIN KERLAP KERLIP PINK MUDA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00324', name: 'KAIN KERLAP KERLIP PINK TUA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00325', name: 'KAIN KERLAP KERLIP UNGU', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00326', name: 'TAPLAK MOTIF KOTAK MERAH PUTIH 200X50', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00327', name: 'TULLE GLITER UNGU', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00328', name: 'TULLE GLITER PINK', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00329', name: 'TULLE GLITER PEACH', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00330', name: 'OSTRICH PINK', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00331', name: 'KAIN UNGU RENDA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00332', name: 'KAIN PINK RENDA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00333', name: 'KAIN PINK MUDA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00334', name: 'KAIN PINK TUA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00335', name: 'TAPLAK PINK RENDA', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00336', name: 'TAPLAK PINK', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: 'SK4', minStock: 1 },
            { id: 'D00337', name: 'TAPLAK GOLD', category: 'Dekorasi', stock: 18, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00338', name: 'LIST RENDA GOLD', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00339', name: 'LIST TAPLAK MEJA GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00340', name: 'TAPLAK KARUNG GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00341', name: 'KAIN GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00342', name: 'KAIN COKELAT CHIFON', category: 'Dekorasi', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00343', name: 'KAIN HIJAU MUDA CHIFON', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00344', name: 'KAIN TOSCA CHIFON', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00345', name: 'KAIN HIJAU', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00346', name: 'TAPLAK MATCHA', category: 'Dekorasi', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: 'SK5', minStock: 1 },
            { id: 'D00347', name: 'TAPLAK BIRU BESAR RENDA GOLD', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00348', name: 'TAPLAK BIRU LAUT', category: 'Dekorasi', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00349', name: 'KAIN BIRU LAUT', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00350', name: 'KAIN BIRU MUDA CHIFON', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00351', name: 'KAIN BIRU MUDA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00352', name: 'TAPLAK MOTIF KOTAK BIRU PUTIH', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00353', name: 'TAPLAK BIRU MUDA RENDA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00354', name: 'TAPLAK BIRU TUA', category: 'Dekorasi', stock: 23, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6,SK2', minStock: 1 },
            { id: 'D00355', name: 'TAPLAK HITAM', category: 'Dekorasi', stock: 17, unit: 'pcs', location: 'Gudang Utama', rack: 'SK6', minStock: 1 },
            { id: 'D00356', name: 'KAIN MAROON BLUDRU', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK7', minStock: 1 },
            { id: 'D00357', name: 'KAIN MAROON', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'SK7', minStock: 1 },
            { id: 'D00358', name: 'HORDENG MERAH', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK7', minStock: 1 },
            { id: 'D00359', name: 'KAIN MIX', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'SK8', minStock: 1 },
            { id: 'D00360', name: 'KAIN IVORY', category: 'Dekorasi', stock: 7, unit: 'pcs', location: 'Gudang Utama', rack: 'SK8', minStock: 1 },
            { id: 'D00361', name: 'BOLA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '42 (EX AXIS)', minStock: 1 },
            { id: 'D00362', name: 'POMPA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '42 (EX AXIS)', minStock: 1 },
            { id: 'D00363', name: 'PLUIT', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '42 (EX AXIS)', minStock: 1 },
            { id: 'D00364', name: 'HORN', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '42 (EX AXIS)', minStock: 1 },
            { id: 'D00365', name: 'SCORE BOARD', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '42 (EX AXIS)', minStock: 1 },
            { id: 'D00366', name: 'GELAS CRYSTAL', category: 'Dekorasi', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '43', minStock: 1 },
            { id: 'D00367', name: 'SLATE COASTER', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '43', minStock: 1 },
            { id: 'D00368', name: 'SANGO TEA CUP SET', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '44', minStock: 1 },
            { id: 'D00369', name: 'HIASAN PIRING GANTUNG MOTIF BUNGA', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '44', minStock: 1 },
            { id: 'D00370', name: 'STYVES BOWL CUP', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '44', minStock: 1 },
            { id: 'D00371', name: 'INFORMA GLASS JAR 950ML', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '45', minStock: 1 },
            { id: 'D00372', name: 'PORSELEN 230CC CUP D SAVER', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '45', minStock: 1 },
            { id: 'D00373', name: 'CANDY BOWL', category: 'Dekorasi', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '45', minStock: 1 },
            { id: 'D00374', name: 'CAKE STAND', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '45', minStock: 1 },
            { id: 'D00375', name: 'ROUND JAR WITH LID 3000ML', category: 'Dekorasi', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '45', minStock: 1 },
            { id: 'D00376', name: 'TOPENG HITAM WOMEN', category: 'Dekorasi', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '46', minStock: 1 },
            { id: 'D00377', name: 'TOPENG HITAM WOMEN W/TIANG', category: 'Dekorasi', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '46', minStock: 1 },
            { id: 'D00378', name: 'DUMMY BOOK', category: 'Dekorasi', stock: 17, unit: 'pcs', location: 'Gudang Utama', rack: '46', minStock: 1 },
            { id: 'D00379', name: 'DUMMY METERAN LISTRIK', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '46', minStock: 1 },
            { id: 'D00380', name: 'TOPI BAJAK LAUT', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '46', minStock: 1 },
            { id: 'D00381', name: 'PIRING LIST EMAS DIA9', category: 'Dekorasi', stock: 18, unit: 'pcs', location: 'Gudang Utama', rack: '47', minStock: 1 },
            { id: 'D00382', name: 'PHOTO FRAME 5R MOTIF KERANG', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '47', minStock: 1 },
            { id: 'D00383', name: 'PHOTO FRAME 4R UKIR GOLD', category: 'Dekorasi', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '47', minStock: 1 },
            { id: 'D00384', name: 'PHOTO FRAME 5R BROWN', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '47', minStock: 1 },
            { id: 'D00385', name: 'STEROFOAM LAMPU ALADIN', category: 'Dekorasi', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00386', name: 'VAS BUNGA KERANJANG', category: 'Dekorasi', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00387', name: 'ORNAMEN SALJU', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00388', name: 'FLAMELESS CANDLE INFORMA', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '29', minStock: 1 },
            { id: 'D00389', name: 'BAKI KAYU 30X20X5', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00390', name: 'LETTER BOARD 30X30', category: 'Dekorasi', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'D00391', name: 'ACRYLIC SIGNAGE DISPLAY A6', category: 'Dekorasi', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0001', name: 'PRINTER EPSON L3210', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0002', name: 'PRINTER EPSON L3211', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0003', name: 'PRINTER BROTHER', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0004', name: 'LG TV 55"', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0005', name: 'LG TV 43"', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0006', name: 'MATADOR TV', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0007', name: 'IGLOO COOLER BOX GREY', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0008', name: 'HT', category: 'Elektronik', stock: 15, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0009', name: 'BACKDROP PORTABLE', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0010', name: 'TROLEY DORONG', category: 'Elektronik', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0011', name: 'TROLEY TARIK', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0012', name: 'SCAFOLDING', category: 'Elektronik', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0013', name: 'TRIPOD SIGNAGE', category: 'Elektronik', stock: 39, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0014', name: 'RAM', category: 'Elektronik', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0015', name: 'KRISBOW CLEANING SET', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0016', name: 'MISTYFAN', category: 'Elektronik', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0017', name: 'STANDING TV', category: 'Elektronik', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0018', name: 'VACUUM CLEANER AERO21 INOX', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0019', name: 'VACUUM HAN RIVER', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0020', name: 'KULKAS MINIBAR GEA 89L GMB91', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0021', name: 'KULKAS MINIBAR GEA 46L GMB50', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0022', name: 'NOISE SPEAKER 12\'', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0023', name: 'SAFETY HELMET BIRU', category: 'Elektronik', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0024', name: 'SAFETY HELMET KUNING', category: 'Elektronik', stock: 41, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0025', name: 'SAFETY HELMET ORANGE', category: 'Elektronik', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0026', name: 'SAFETY ROMPI HIJAU', category: 'Elektronik', stock: 63, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0027', name: 'SAFETY ROMPI KUNING', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0028', name: 'SAFETY ROMPI ORANGE', category: 'Elektronik', stock: 35, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0029', name: 'BODY HARNES', category: 'Elektronik', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0030', name: 'SARUNG TANGAN LISTRIK', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0031', name: 'ACRYLIC A5 2 KAKI', category: 'Elektronik', stock: 18, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0032', name: 'ACRYLIC A5 1 KAKI', category: 'Elektronik', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0033', name: 'ACRYLIC A4 2 KAKI', category: 'Elektronik', stock: 16, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0034', name: 'ACRYLIC A4 1 KAKI', category: 'Elektronik', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0035', name: 'STAND ACRYLIC 9X7', category: 'Elektronik', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0036', name: 'STAND ACRYLIC SIM/CARD HOLDER', category: 'Elektronik', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0037', name: 'ACRYLIC BROSUR HOLDER BESAR', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0038', name: 'ACRYLIC BROSUR HOLDER KECIL', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0039', name: 'ACRYLIC SEGITIGA 7X30', category: 'Elektronik', stock: 13, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0040', name: 'ACRYLIC SEGITIGA 7X25', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0041', name: 'ACRYLIC SEGITGA 8X30', category: 'Elektronik', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0042', name: 'ACRYLIC TABLE A6', category: 'Elektronik', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '49', minStock: 1 },
            { id: 'ES0043', name: 'SANDBAG', category: 'Elektronik', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0044', name: 'SETERIKA UAP ARIETE', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0045', name: 'MEASURING  TAPE', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0046', name: 'PAYUNG', category: 'Elektronik', stock: 14, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0047', name: 'TRASHBIN STORA 10L', category: 'Elektronik', stock: 12, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0048', name: 'TEKO LISTRIK', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0049', name: 'TOTE BAG TERPAL', category: 'Elektronik', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0050', name: 'KASUR ANGIN BESTWAY', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0051', name: 'BENDERA MERAH PUTIH 136X90', category: 'Elektronik', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0052', name: 'PATAKA', category: 'Elektronik', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0053', name: 'SAFETY SHOES', category: 'Elektronik', stock: 31, unit: 'pcs', location: 'Gudang Utama', rack: '48', minStock: 1 },
            { id: 'ES0054', name: 'SAFETY ROMPI BIRU MUDA', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0055', name: 'SAFETY ROMPI BIRU TUA', category: 'Elektronik', stock: 0, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0056', name: 'METERAN ROLL', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0057', name: 'POINTER LASER', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0058', name: 'KUNCI SHOCK', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0059', name: 'SAFETY HELMET PUTIH', category: 'Elektronik', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'ES0060', name: 'BAKI KAYU IKEA PUTH 57X37', category: 'Elektronik', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0001', name: 'BARSTOOL', category: 'Furniture', stock: 6, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0002', name: 'MEJA LESEHAN', category: 'Furniture', stock: 10, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0003', name: 'PUFF CHAIR BIRU', category: 'Furniture', stock: 5, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0004', name: 'PUFF CHAIR PUTIH', category: 'Furniture', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0005', name: 'RAK SUSUN 40X60X160', category: 'Furniture', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0006', name: 'RAK BROSUR SUSUN PUTIH', category: 'Furniture', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0007', name: 'COFFEE TABLE PUTIH', category: 'Furniture', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0008', name: 'PODIUM ACRYLIC', category: 'Furniture', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0009', name: 'MEJA BAKSO', category: 'Furniture', stock: 8, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0010', name: 'KURSI BAKSO', category: 'Furniture', stock: 20, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'FU0011', name: 'BENCH MERAH', category: 'Furniture', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: '', minStock: 1 },
            { id: 'SS0001', name: 'KABEL AKAI TO XLR 25M', category: 'Sound System', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'SS0002', name: 'KABEK AKAI TO XLR 2m', category: 'Sound System', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'SS0003', name: 'KABEL JACK 3.5', category: 'Sound System', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'SS0004', name: 'KABEL XLR TO XLR 5m', category: 'Sound System', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'SS0005', name: 'KABEL XLR TO XLR 10m', category: 'Sound System', stock: 4, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'SS0006', name: 'KABEL XLR MALE TO JACK 3.5 1m', category: 'Sound System', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'SS0007', name: 'QA MIC WIREESS', category: 'Sound System', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT SOUND', minStock: 1 },
            { id: 'MM0001', name: 'KABEL LAN 5M', category: 'Multimedia', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0002', name: 'WIRELES RECEIVER TRANSMITTER', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0003', name: 'KABEL FO AUDIO NYK 1.5M', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0004', name: 'HEADSET LOGITEC', category: 'Multimedia', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0005', name: 'BLUETOOTH DONGLE 5.0', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0006', name: 'HEADPHONE SPLITTER 3.5MM 6PORT', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0007', name: 'ORBIT TELKOMSEL', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0008', name: 'USB 6 PORT', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0009', name: 'KABEL 3.5MM TO RCA STEREO 1M', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0010', name: 'TOSLINK CABLE', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0011', name: 'HDMI VIDEO CAPTURE', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0012', name: 'AUDIO CONVERTER', category: 'Multimedia', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0013', name: 'UGREEN ALL IN ONE DATA CABLE SET', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0014', name: 'SANDISK 128GB', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0015', name: 'USB 4GB', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0016', name: 'DONGLE HDMI', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 },
            { id: 'MM0017', name: 'KABEL HDMI 25M BAFO', category: 'Multimedia', stock: 3, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0018', name: 'KABEL HDMI 3M VENTION', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0019', name: 'KABEL HDMI 1M VENTION', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0020', name: 'SPLITTER HDMI 4PORT', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0021', name: 'EXTENDER HDMI TO LAN', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0022', name: 'USB EXTENTION 5M BIRU', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0023', name: 'USB EXTENTION 10M BIRU', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0024', name: 'USB EXTENTION 10M GREY', category: 'Multimedia', stock: 1, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT HDMI', minStock: 1 },
            { id: 'MM0025', name: 'WEB CAM MTECH', category: 'Multimedia', stock: 2, unit: 'pcs', location: 'Gudang Utama', rack: 'CONT MULMED', minStock: 1 }
];
function seedDB() {
    return {
        users: [
            { id: 'u-staff', username: 'staff', password: 'staff123', name: 'Budi Santoso', role: 'staff' },
            { id: 'u-admin', username: 'admin', password: 'admin123', name: 'Sari Wijaya', role: 'admin' },
            { id: 'u-mgr', username: 'manajemen', password: 'manajemen123', name: 'Hendra Kusuma', role: 'manajemen' }
        ],
        items: JSON.parse(JSON.stringify(IMPORTED_ITEMS)),
        loans: [],
        damages: [],
        purchases: [],
        requests: [],
        lost: [],
        itemsDataVersion: ITEMS_DATA_VERSION
    };
}
function getDB() {
    let raw = localStorage.getItem(DB_KEY);
    if (!raw) {
        const seeded = seedDB();
        localStorage.setItem(DB_KEY, JSON.stringify(seeded));
        return seeded;
    }
    let db;
    try { db = JSON.parse(raw); } catch (e) { const seeded = seedDB(); localStorage.setItem(DB_KEY, JSON.stringify(seeded)); return seeded; }
    // Migrasi: perbarui daftar barang ke data import terbaru jika versi berbeda,
    // tanpa menghapus riwayat peminjaman/pembelian/dll yang sudah ada.
    if (db.itemsDataVersion !== ITEMS_DATA_VERSION) {
        db.items = JSON.parse(JSON.stringify(IMPORTED_ITEMS));
        db.itemsDataVersion = ITEMS_DATA_VERSION;
        localStorage.setItem(DB_KEY, JSON.stringify(db));
    }
    return db;
}
function saveDB(db) {
    try {
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        return true;
    } catch (e) {
        toast('Penyimpanan penuh (foto invoice terlalu besar/banyak). Data terakhir gagal disimpan.', 'danger');
        return false;
    }
}

/* ---------- Session ---------- */
function getSession() {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY)); } catch (e) { return null; }
}
function setSession(user) {
    localStorage.setItem(SESSION_KEY, JSON.stringify({ id: user.id, username: user.username, name: user.name, role: user.role }));
}
function clearSession() { localStorage.removeItem(SESSION_KEY); }
function currentUser() { return getSession(); }
function hasRole(...roles) {
    const u = currentUser();
    return u && roles.includes(u.role);
}

/* ---------- Toast ---------- */
function toast(msg, type = 'success') {
    const container = document.getElementById('toast-container');
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    const icon = type === 'success' ? 'check-circle' : type === 'danger' ? 'x-circle' : 'alert-triangle';
    el.innerHTML = `<i data-lucide="${icon}"></i><span>${escapeHtml(msg)}</span>`;
    container.appendChild(el);
    if (window.lucide) lucide.createIcons();
    setTimeout(() => { el.remove(); }, 3800);
}

/* ---------- Modal ---------- */
function openModal(title, bodyHtml) {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-body').innerHTML = bodyHtml;
    document.getElementById('modal').classList.remove('hidden');
    if (window.lucide) lucide.createIcons();
}
function closeModal() {
    document.getElementById('modal').classList.add('hidden');
    document.getElementById('modal-body').innerHTML = '';
}

/* ---------- Navigasi (menu per role) ---------- */
const NAV_CONFIG = [
    { section: 'Umum', items: [
        { id: 'dashboard', label: 'Dashboard', icon: 'layout-dashboard', roles: ['staff', 'admin', 'manajemen'] },
        { id: 'items', label: 'Data Barang', icon: 'package', roles: ['staff', 'admin', 'manajemen'] },
        { id: 'loans', label: 'Peminjaman & Pengembalian', icon: 'repeat', roles: ['staff', 'admin', 'manajemen'] },
        { id: 'requests', label: 'Pengajuan Pembelian', icon: 'file-text', roles: ['staff', 'admin', 'manajemen'] }
    ]},
    { section: 'Gudang (Admin)', items: [
        { id: 'purchases', label: 'Pembelian Barang', icon: 'shopping-cart', roles: ['admin', 'manajemen'] },
        { id: 'damages', label: 'Barang Rusak / MT', icon: 'wrench', roles: ['admin', 'manajemen'] },
        { id: 'lost', label: 'Barang Hilang', icon: 'search-x', roles: ['admin', 'manajemen'] }
    ]},
    { section: 'Manajemen', items: [
        { id: 'users', label: 'Manajemen User', icon: 'users', roles: ['manajemen'] }
    ]}
];
const PAGE_TITLES = {
    dashboard: 'Dashboard', items: 'Data Barang', loans: 'Peminjaman & Pengembalian',
    purchases: 'Pembelian Barang Baru', requests: 'Pengajuan Pembelian', damages: 'Barang Rusak / Maintenance',
    lost: 'Barang Hilang', users: 'Manajemen User'
};
let currentPage = 'dashboard';

function loanActionCount() {
    const db = getDB();
    const u = currentUser();
    if (!u) return 0;
    if (u.role === 'admin') return db.loans.filter(l => l.status === 'menunggu_admin').length;
    if (u.role === 'manajemen') return db.loans.filter(l => l.status === 'menunggu_admin' || l.status === 'menunggu_manajemen').length;
    // staff: informasikan jumlah pengajuan miliknya yang masih menunggu
    return db.loans.filter(l => l.requestedBy === u.name && (l.status === 'menunggu_admin' || l.status === 'menunggu_manajemen')).length;
}
function renderSidebar() {
    const nav = document.getElementById('sidebar-nav');
    const db = getDB();
    const pendingReq = db.requests.filter(r => r.status === 'pending').length;
    let html = '';
    NAV_CONFIG.forEach(sec => {
        const visible = sec.items.filter(it => it.roles.includes(currentUser().role));
        if (!visible.length) return;
        html += `<div class="nav-section-label">${escapeHtml(sec.section)}</div>`;
        visible.forEach(it => {
            let badge = '';
            if (it.id === 'requests' && hasRole('admin', 'manajemen') && pendingReq > 0) badge = `<span class="nav-badge">${pendingReq}</span>`;
            if (it.id === 'loans') {
                const n = loanActionCount();
                if (n > 0) badge = `<span class="nav-badge">${n}</span>`;
            }
            html += `<div class="nav-item ${currentPage === it.id ? 'active' : ''}" onclick="navigateTo('${it.id}')">
                <i data-lucide="${it.icon}"></i><span>${escapeHtml(it.label)}</span>${badge}
            </div>`;
        });
    });
    nav.innerHTML = html;
    if (window.lucide) lucide.createIcons();
}
function navigateTo(pageId) {
    currentPage = pageId;
    document.getElementById('page-title').textContent = PAGE_TITLES[pageId] || '';
    document.getElementById('sidebar').classList.remove('open');
    renderSidebar();
    renderPage(pageId);
}
function renderPage(pageId) {
    const map = {
        dashboard: renderDashboard, items: renderItemsPage, loans: renderLoansPage,
        purchases: renderPurchasesPage, requests: renderRequestsPage, damages: renderDamagesPage,
        lost: renderLostPage, users: renderUsersPage
    };
    const fn = map[pageId] || renderDashboard;
    document.getElementById('content-area').innerHTML = fn();
    if (window.lucide) lucide.createIcons();
    afterRenderHooks[pageId] && afterRenderHooks[pageId]();
}
const afterRenderHooks = {};

/* =========================================================
   DASHBOARD
   ========================================================= */
function renderDashboard() {
    const db = getDB();
    const totalItems = db.items.length;
    const lowStock = db.items.filter(i => i.stock <= i.minStock).length;
    const activeLoans = db.loans.filter(l => l.status === 'dipinjam').length;
    const pendingReq = db.requests.filter(r => r.status === 'pending').length;

    const recent = [
        ...db.loans.map(l => ({ date: l.borrowDate, text: `${l.borrower} meminjam ${l.itemName} (${l.qty})`, icon: 'repeat' })),
        ...db.purchases.map(p => ({ date: p.date, text: `Pembelian baru: ${p.itemName} (${p.qty})`, icon: 'shopping-cart' })),
        ...db.damages.map(d => ({ date: d.date, text: `${d.type === 'mt' ? 'Perlu maintenance' : 'Barang rusak'}: ${d.itemName}`, icon: 'wrench' })),
        ...db.requests.map(r => ({ date: r.date, text: `Pengajuan pembelian: ${r.itemName} oleh ${r.requestedBy}`, icon: 'file-text' })),
        ...db.lost.map(l => ({ date: l.date, text: `Barang hilang dilaporkan: ${l.itemName}`, icon: 'search-x' }))
    ].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 8);

    return `
    <div class="card-grid">
        <div class="stat-card">
            <div class="stat-icon"><i data-lucide="package"></i></div>
            <div><div class="stat-value">${totalItems}</div><div class="stat-label">Jenis Barang</div></div>
        </div>
        <div class="stat-card">
            <div class="stat-icon warning"><i data-lucide="alert-triangle"></i></div>
            <div><div class="stat-value">${lowStock}</div><div class="stat-label">Stok Menipis</div></div>
        </div>
        <div class="stat-card">
            <div class="stat-icon success"><i data-lucide="repeat"></i></div>
            <div><div class="stat-value">${activeLoans}</div><div class="stat-label">Sedang Dipinjam</div></div>
        </div>
        <div class="stat-card">
            <div class="stat-icon purple"><i data-lucide="file-text"></i></div>
            <div><div class="stat-value">${pendingReq}</div><div class="stat-label">Pengajuan Menunggu</div></div>
        </div>
    </div>
    <div class="table-container">
        <div class="table-header"><h3>Aktivitas Terbaru</h3></div>
        <table class="data-table">
            <thead><tr><th>Tanggal</th><th>Aktivitas</th></tr></thead>
            <tbody>
                ${recent.length ? recent.map(r => `<tr><td>${fmtDate(r.date)}</td><td>${escapeHtml(r.text)}</td></tr>`).join('') : '<tr class="empty-row"><td colspan="2">Belum ada aktivitas.</td></tr>'}
            </tbody>
        </table>
    </div>`;
}

/* =========================================================
   DATA BARANG (lokasi & rak terpisah)
   ========================================================= */
function renderItemsPage() {
    const db = getDB();
    const canEdit = hasRole('admin', 'manajemen');
    return `
    <div class="page-header">
        <div><p>Kelola data barang, stok, lokasi, dan rak penyimpanan.</p></div>
        ${canEdit ? `<button class="btn btn-primary" onclick="openItemForm()"><i data-lucide="plus"></i>Tambah Barang</button>` : ''}
    </div>
    <div class="filter-bar">
        <input type="text" id="item-search" placeholder="Cari nama barang..." oninput="filterItemsTable()">
    </div>
    <div class="table-container">
        <table class="data-table">
            <thead><tr>
                <th>Nama Barang</th><th>Kategori</th><th>Stok</th><th>Satuan</th><th>Lokasi</th><th>Rak</th><th>Status</th>${canEdit ? '<th>Aksi</th>' : ''}
            </tr></thead>
            <tbody id="items-tbody">
                ${itemsTableRows(db.items, canEdit)}
            </tbody>
        </table>
    </div>`;
}
function itemsTableRows(items, canEdit) {
    if (!items.length) return `<tr class="empty-row"><td colspan="8">Belum ada data barang.</td></tr>`;
    return items.map(it => {
        const low = it.stock <= it.minStock;
        return `<tr>
            <td>${escapeHtml(it.name)}</td>
            <td>${escapeHtml(it.category || '-')}</td>
            <td>${it.stock}</td>
            <td>${escapeHtml(it.unit || '-')}</td>
            <td>${escapeHtml(it.location || '-')}</td>
            <td>${escapeHtml(it.rack || '-')}</td>
            <td>${low ? '<span class="badge badge-danger">Stok Menipis</span>' : '<span class="badge badge-success">Aman</span>'}</td>
            ${canEdit ? `<td class="actions">
                <button class="btn-icon" title="Edit" onclick="openItemForm('${it.id}')"><i data-lucide="pencil"></i></button>
                <button class="btn-icon text-danger" title="Hapus" onclick="deleteItem('${it.id}')"><i data-lucide="trash-2"></i></button>
            </td>` : ''}
        </tr>`;
    }).join('');
}
function filterItemsTable() {
    const q = document.getElementById('item-search').value.toLowerCase();
    const db = getDB();
    const filtered = db.items.filter(i => i.name.toLowerCase().includes(q));
    document.getElementById('items-tbody').innerHTML = itemsTableRows(filtered, hasRole('admin', 'manajemen'));
    if (window.lucide) lucide.createIcons();
}
function openItemForm(itemId) {
    const db = getDB();
    const item = itemId ? db.items.find(i => i.id === itemId) : null;
    openModal(item ? 'Edit Barang' : 'Tambah Barang Baru', `
        <form id="item-form" onsubmit="submitItemForm(event, '${itemId || ''}')">
            <div class="form-group"><label>Nama Barang</label><input type="text" id="f-name" value="${item ? escapeHtml(item.name) : ''}" required></div>
            <div class="form-row">
                <div class="form-group"><label>Kategori</label><input type="text" id="f-category" value="${item ? escapeHtml(item.category || '') : ''}"></div>
                <div class="form-group"><label>Satuan</label><input type="text" id="f-unit" placeholder="pcs / unit / roll" value="${item ? escapeHtml(item.unit || '') : ''}"></div>
            </div>
            <div class="form-row">
                <div class="form-group"><label>Stok Saat Ini</label><input type="number" id="f-stock" min="0" value="${item ? item.stock : 0}" required></div>
                <div class="form-group"><label>Stok Minimum</label><input type="number" id="f-minstock" min="0" value="${item ? item.minStock : 1}" required></div>
            </div>
            <div class="form-row">
                <div class="form-group"><label>Lokasi (Gudang)</label><input type="text" id="f-location" placeholder="cth: Gudang A" value="${item ? escapeHtml(item.location || '') : ''}" required></div>
                <div class="form-group"><label>Rak</label><input type="text" id="f-rack" placeholder="cth: Rak 2-B" value="${item ? escapeHtml(item.rack || '') : ''}" required></div>
            </div>
            <button type="submit" class="btn btn-primary btn-block">Simpan</button>
        </form>
    `);
}
function submitItemForm(e, itemId) {
    e.preventDefault();
    const db = getDB();
    const data = {
        name: document.getElementById('f-name').value.trim(),
        category: document.getElementById('f-category').value.trim(),
        unit: document.getElementById('f-unit').value.trim(),
        stock: Number(document.getElementById('f-stock').value),
        minStock: Number(document.getElementById('f-minstock').value),
        location: document.getElementById('f-location').value.trim(),
        rack: document.getElementById('f-rack').value.trim()
    };
    if (itemId) {
        const idx = db.items.findIndex(i => i.id === itemId);
        db.items[idx] = { ...db.items[idx], ...data };
        toast('Barang berhasil diperbarui.');
    } else {
        db.items.push({ id: uid('itm'), ...data });
        toast('Barang baru berhasil ditambahkan.');
    }
    saveDB(db);
    closeModal();
    renderPage('items');
    renderSidebar();
}
function deleteItem(itemId) {
    if (!confirm('Hapus barang ini? Tindakan tidak dapat dibatalkan.')) return;
    const db = getDB();
    db.items = db.items.filter(i => i.id !== itemId);
    saveDB(db);
    renderPage('items');
    toast('Barang dihapus.', 'warning');
}

/* =========================================================
   PEMINJAMAN & PENGEMBALIAN (+ form Barang Rusak/MT khusus admin)
   ========================================================= */
let loanTab = 'pinjam';
function renderLoansPage() {
    const db = getDB();
    const approvalCount = db.loans.filter(l => l.status === 'menunggu_admin' || l.status === 'menunggu_manajemen').length;
    return `
    <div class="tabs">
        <button class="tab-btn ${loanTab === 'pinjam' ? 'active' : ''}" onclick="switchLoanTab('pinjam')">Ajukan Peminjaman</button>
        <button class="tab-btn ${loanTab === 'persetujuan' ? 'active' : ''}" onclick="switchLoanTab('persetujuan')">Menunggu Persetujuan${approvalCount ? ` (${approvalCount})` : ''}</button>
        <button class="tab-btn ${loanTab === 'aktif' ? 'active' : ''}" onclick="switchLoanTab('aktif')">Sedang Dipinjam</button>
        <button class="tab-btn ${loanTab === 'riwayat' ? 'active' : ''}" onclick="switchLoanTab('riwayat')">Riwayat</button>
    </div>
    <div id="loans-tab-content">${renderLoanTabContent()}</div>`;
}
function switchLoanTab(tab) { loanTab = tab; renderPage('loans'); }
function renderLoanTabContent() {
    const db = getDB();
    if (loanTab === 'pinjam') {
        const options = db.items.filter(i => i.stock > 0).map(i => `<option value="${i.id}">${escapeHtml(i.name)} (stok: ${i.stock} ${escapeHtml(i.unit || '')})</option>`).join('');
        return `
        <div class="table-container">
            <div class="modal-body">
                <form id="loan-form" onsubmit="submitLoanForm(event)">
                    <div class="form-group"><label>Pilih Barang</label>
                        <select id="loan-item" required>${options || '<option value="">Tidak ada barang tersedia</option>'}</select>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Jumlah</label><input type="number" id="loan-qty" min="1" value="1" required></div>
                        <div class="form-group"><label>Nama Peminjam</label><input type="text" id="loan-borrower" value="${escapeHtml(currentUser().name)}" required></div>
                    </div>
                    <div class="form-row">
                        <div class="form-group"><label>Tanggal Pinjam</label><input type="date" id="loan-date" value="${todayStr()}" required></div>
                        <div class="form-group"><label>Rencana Kembali</label><input type="date" id="loan-return-plan" required></div>
                    </div>
                    <div class="form-group"><label>Keperluan / Catatan</label><textarea id="loan-notes" placeholder="cth: untuk perbaikan mesin di lokasi X"></textarea></div>
                    <p class="text-muted" style="margin-bottom:12px;">Pengajuan ini akan diproses menunggu persetujuan Admin, lalu Manajemen, sebelum barang boleh diambil.</p>
                    <button type="submit" class="btn btn-primary">Ajukan Peminjaman</button>
                </form>
            </div>
        </div>`;
    }
    if (loanTab === 'persetujuan') {
        const u = currentUser();
        let list = db.loans.filter(l => l.status === 'menunggu_admin' || l.status === 'menunggu_manajemen');
        if (u.role === 'staff') list = list.filter(l => l.requestedBy === u.name);
        list = list.sort((a, b) => new Date(b.borrowDate) - new Date(a.borrowDate));
        return `
        <div class="table-container">
            <table class="data-table">
                <thead><tr><th>Barang</th><th>Jml</th><th>Peminjam</th><th>Diajukan Oleh</th><th>Tgl Pinjam</th><th>Status</th><th>Aksi</th></tr></thead>
                <tbody>
                    ${list.length ? list.map(l => `<tr>
                        <td>${escapeHtml(l.itemName)}</td><td>${l.qty}</td><td>${escapeHtml(l.borrower)}</td>
                        <td>${escapeHtml(l.requestedBy)}</td><td>${fmtDate(l.borrowDate)}</td>
                        <td>${loanStatusBadge(l.status)}</td>
                        <td class="actions">
                            ${hasRole('admin', 'manajemen') && l.status === 'menunggu_admin' ? `
                                <button class="btn btn-sm btn-success" onclick="approveLoanStage('${l.id}','admin')"><i data-lucide="check"></i>Setujui (Admin)</button>
                                <button class="btn btn-sm btn-danger" onclick="rejectLoan('${l.id}')"><i data-lucide="x"></i>Tolak</button>
                            ` : ''}
                            ${hasRole('manajemen') && l.status === 'menunggu_manajemen' ? `
                                <button class="btn btn-sm btn-success" onclick="approveLoanStage('${l.id}','manajemen')"><i data-lucide="check"></i>Setujui (Manajemen)</button>
                                <button class="btn btn-sm btn-danger" onclick="rejectLoan('${l.id}')"><i data-lucide="x"></i>Tolak</button>
                            ` : ''}
                            ${!hasRole('admin', 'manajemen') || (l.status === 'menunggu_manajemen' && !hasRole('manajemen')) ? (hasRole('admin','manajemen') ? '' : '<span class="text-muted">Menunggu proses</span>') : ''}
                        </td>
                    </tr>`).join('') : `<tr class="empty-row"><td colspan="7">Tidak ada pengajuan yang menunggu persetujuan.</td></tr>`}
                </tbody>
            </table>
        </div>`;
    }
    if (loanTab === 'aktif') {
        const active = db.loans.filter(l => l.status === 'dipinjam');
        return `
        <div class="table-container">
            <table class="data-table">
                <thead><tr><th>Barang</th><th>Jml</th><th>Peminjam</th><th>Tgl Pinjam</th><th>Rencana Kembali</th><th>Disetujui Oleh</th>${hasRole('admin', 'manajemen') ? '<th>Aksi</th>' : ''}</tr></thead>
                <tbody>
                    ${active.length ? active.map(l => `<tr>
                        <td>${escapeHtml(l.itemName)}</td><td>${l.qty}</td><td>${escapeHtml(l.borrower)}</td>
                        <td>${fmtDate(l.borrowDate)}</td><td>${fmtDate(l.returnPlan)}</td>
                        <td>${escapeHtml(l.adminApprovedBy || '-')} &amp; ${escapeHtml(l.managementApprovedBy || '-')}</td>
                        ${hasRole('admin', 'manajemen') ? `<td><button class="btn btn-sm btn-success" onclick="openReturnForm('${l.id}')"><i data-lucide="corner-down-left"></i>Proses Pengembalian</button></td>` : ''}
                    </tr>`).join('') : `<tr class="empty-row"><td colspan="7">Tidak ada peminjaman aktif.</td></tr>`}
                </tbody>
            </table>
        </div>`;
    }
    // riwayat
    const done = db.loans.filter(l => l.status === 'dikembalikan' || l.status === 'ditolak')
        .sort((a, b) => new Date(b.returnDate || b.rejectedDate || b.borrowDate) - new Date(a.returnDate || a.rejectedDate || a.borrowDate));
    return `
    <div class="table-container">
        <table class="data-table">
            <thead><tr><th>Barang</th><th>Jml</th><th>Peminjam</th><th>Tgl Pinjam</th><th>Status</th><th>Detail</th></tr></thead>
            <tbody>
                ${done.length ? done.map(l => `<tr>
                    <td>${escapeHtml(l.itemName)}</td><td>${l.qty}</td><td>${escapeHtml(l.borrower)}</td>
                    <td>${fmtDate(l.borrowDate)}</td>
                    <td>${l.status === 'ditolak' ? '<span class="badge badge-danger">Ditolak</span>' : conditionBadge(l.condition)}</td>
                    <td>${l.status === 'ditolak' ? `Ditolak oleh ${escapeHtml(l.rejectedBy || '-')}${l.rejectReason ? ': ' + escapeHtml(l.rejectReason) : ''}` : `Kembali: ${fmtDate(l.returnDate)}`}</td>
                </tr>`).join('') : `<tr class="empty-row"><td colspan="6">Belum ada riwayat.</td></tr>`}
            </tbody>
        </table>
    </div>`;
}
function loanStatusBadge(status) {
    if (status === 'menunggu_admin') return '<span class="badge badge-warning">Menunggu Admin</span>';
    if (status === 'menunggu_manajemen') return '<span class="badge badge-info">Menunggu Manajemen</span>';
    if (status === 'dipinjam') return '<span class="badge badge-success">Aktif</span>';
    if (status === 'ditolak') return '<span class="badge badge-danger">Ditolak</span>';
    return '-';
}
function conditionBadge(cond) {
    if (cond === 'baik') return '<span class="badge badge-success">Baik</span>';
    if (cond === 'rusak') return '<span class="badge badge-danger">Rusak</span>';
    if (cond === 'mt') return '<span class="badge badge-warning">Perlu Maintenance</span>';
    return '-';
}
function submitLoanForm(e) {
    e.preventDefault();
    const db = getDB();
    const itemId = document.getElementById('loan-item').value;
    const item = db.items.find(i => i.id === itemId);
    const qty = Number(document.getElementById('loan-qty').value);
    if (!item || qty < 1 || qty > item.stock) { toast('Jumlah melebihi stok tersedia.', 'danger'); return; }
    // Stok BELUM dikurangi di sini - baru dikurangi setelah disetujui Admin & Manajemen
    db.loans.push({
        id: uid('loan'), itemId, itemName: item.name, qty,
        borrower: document.getElementById('loan-borrower').value.trim(),
        borrowDate: document.getElementById('loan-date').value,
        returnPlan: document.getElementById('loan-return-plan').value,
        notes: document.getElementById('loan-notes').value.trim(),
        status: 'menunggu_admin', requestedBy: currentUser().name
    });
    saveDB(db);
    toast('Peminjaman berhasil diajukan, menunggu persetujuan Admin.');
    switchLoanTab('persetujuan');
    renderSidebar();
}
// Persetujuan bertingkat: Admin dulu, baru Manajemen. Stok baru dipotong setelah keduanya setuju.
function approveLoanStage(loanId, stage) {
    const db = getDB();
    const loan = db.loans.find(l => l.id === loanId);
    if (!loan) return;
    if (stage === 'admin') {
        if (loan.status !== 'menunggu_admin') return;
        loan.status = 'menunggu_manajemen';
        loan.adminApprovedBy = currentUser().name;
        loan.adminApprovedDate = todayStr();
        saveDB(db);
        toast('Peminjaman disetujui Admin, diteruskan ke Manajemen.');
    } else if (stage === 'manajemen') {
        if (loan.status !== 'menunggu_manajemen') return;
        const item = db.items.find(i => i.id === loan.itemId);
        if (!item || item.stock < loan.qty) {
            toast('Stok tidak lagi mencukupi. Peminjaman tidak dapat disetujui.', 'danger');
            return;
        }
        item.stock -= loan.qty;
        loan.status = 'dipinjam';
        loan.managementApprovedBy = currentUser().name;
        loan.managementApprovedDate = todayStr();
        saveDB(db);
        toast('Peminjaman disetujui Manajemen. Stok diperbarui, barang siap diambil.');
    }
    renderPage('loans');
    renderSidebar();
}
function rejectLoan(loanId) {
    const db = getDB();
    const loan = db.loans.find(l => l.id === loanId);
    if (!loan) return;
    const reason = prompt('Alasan penolakan (opsional):', '') || '';
    loan.status = 'ditolak';
    loan.rejectedBy = currentUser().name;
    loan.rejectedDate = todayStr();
    loan.rejectReason = reason.trim();
    saveDB(db);
    toast('Pengajuan peminjaman ditolak.', 'warning');
    renderPage('loans');
    renderSidebar();
}
// Form khusus admin saat pengembalian: pilih kondisi barang; jika Rusak/MT -> munculkan detail kerusakan
function openReturnForm(loanId) {
    openModal('Proses Pengembalian Barang (Admin)', `
        <form id="return-form" onsubmit="submitReturnForm(event, '${loanId}')">
            <div class="form-group"><label>Kondisi Barang Saat Dikembalikan</label>
                <select id="return-condition" onchange="toggleDamageFields()" required>
                    <option value="baik">Baik (kembali ke stok)</option>
                    <option value="rusak">Rusak</option>
                    <option value="mt">Perlu Maintenance (MT)</option>
                </select>
            </div>
            <div id="damage-fields" class="hidden">
                <div class="form-group"><label>Deskripsi Kerusakan / Kebutuhan MT</label>
                    <textarea id="damage-desc" placeholder="Jelaskan kondisi kerusakan barang..."></textarea>
                </div>
            </div>
            <div class="form-group"><label>Catatan Pengembalian (opsional)</label><textarea id="return-notes"></textarea></div>
            <button type="submit" class="btn btn-primary">Konfirmasi Pengembalian</button>
        </form>
    `);
}
function toggleDamageFields() {
    const val = document.getElementById('return-condition').value;
    document.getElementById('damage-fields').classList.toggle('hidden', val === 'baik');
}
function submitReturnForm(e, loanId) {
    e.preventDefault();
    const db = getDB();
    const loan = db.loans.find(l => l.id === loanId);
    const condition = document.getElementById('return-condition').value;
    loan.status = 'dikembalikan';
    loan.returnDate = todayStr();
    loan.condition = condition;
    loan.returnNotes = document.getElementById('return-notes').value.trim();
    loan.processedBy = currentUser().name;

    if (condition === 'baik') {
        const item = db.items.find(i => i.id === loan.itemId);
        if (item) item.stock += loan.qty;
    } else {
        // Rusak / MT -> tidak kembali ke stok baik, dicatat di modul Barang Rusak/MT
        db.damages.push({
            id: uid('dmg'), itemId: loan.itemId, itemName: loan.itemName, qty: loan.qty,
            type: condition, description: document.getElementById('damage-desc').value.trim(),
            source: 'pengembalian', loanId: loan.id,
            reportedBy: currentUser().name, date: todayStr(), status: 'menunggu'
        });
    }
    saveDB(db);
    closeModal();
    toast(condition === 'baik' ? 'Barang dikembalikan & stok diperbarui.' : 'Barang dikembalikan & dicatat ke Barang Rusak/MT.', condition === 'baik' ? 'success' : 'warning');
    switchLoanTab('aktif');
    renderSidebar();
}

/* =========================================================
   PEMBELIAN BARANG BARU (stok masuk + invoice foto)
   ========================================================= */
function renderPurchasesPage() {
    if (!hasRole('admin', 'manajemen')) return `<p class="text-muted">Anda tidak memiliki akses ke halaman ini.</p>`;
    const db = getDB();
    const rows = [...db.purchases].sort((a, b) => new Date(b.date) - new Date(a.date));
    return `
    <div class="page-header">
        <p>Catat pembelian barang baru untuk menambah stok, lengkap dengan foto invoice.</p>
        <button class="btn btn-primary" onclick="openPurchaseForm()"><i data-lucide="plus"></i>Catat Pembelian</button>
    </div>
    <div class="table-container">
        <table class="data-table">
            <thead><tr><th>Tanggal</th><th>Barang</th><th>Jml</th><th>Harga Satuan</th><th>Total</th><th>Supplier</th><th>Invoice</th></tr></thead>
            <tbody>
                ${rows.length ? rows.map(p => `<tr>
                    <td>${fmtDate(p.date)}</td><td>${escapeHtml(p.itemName)}</td><td>${p.qty}</td>
                    <td>${fmtRupiah(p.price)}</td><td>${fmtRupiah(p.price * p.qty)}</td><td>${escapeHtml(p.supplier || '-')}</td>
                    <td>${p.invoicePhoto ? `<img src="${p.invoicePhoto}" class="thumb" onclick="viewInvoice('${p.id}')" title="Lihat invoice">` : '-'}</td>
                </tr>`).join('') : `<tr class="empty-row"><td colspan="7">Belum ada data pembelian.</td></tr>`}
            </tbody>
        </table>
    </div>`;
}
function openPurchaseForm() {
    const db = getDB();
    const options = db.items.map(i => `<option value="${i.id}">${escapeHtml(i.name)}</option>`).join('');
    openModal('Catat Pembelian Barang Baru', `
        <form id="purchase-form" onsubmit="submitPurchaseForm(event)">
            <div class="form-group"><label>Barang</label>
                <select id="p-item" required>
                    <option value="__new__">+ Barang baru (belum ada di data)</option>
                    ${options}
                </select>
            </div>
            <div class="form-group hidden" id="p-newname-group"><label>Nama Barang Baru</label><input type="text" id="p-newname"></div>
            <div class="form-row">
                <div class="form-group"><label>Jumlah</label><input type="number" id="p-qty" min="1" value="1" required></div>
                <div class="form-group"><label>Harga Satuan (Rp)</label><input type="number" id="p-price" min="0" value="0" required></div>
            </div>
            <div class="form-row">
                <div class="form-group"><label>Supplier</label><input type="text" id="p-supplier"></div>
                <div class="form-group"><label>Tanggal</label><input type="date" id="p-date" value="${todayStr()}" required></div>
            </div>
            <div class="form-group"><label>Foto Invoice</label>
                <div class="upload-box" onclick="document.getElementById('p-invoice-file').click()">
                    <i data-lucide="camera"></i><div>Klik untuk unggah foto invoice</div>
                </div>
                <input type="file" id="p-invoice-file" accept="image/*" class="hidden" onchange="previewInvoice(event)">
                <img id="p-invoice-preview" class="upload-preview hidden">
            </div>
            <button type="submit" class="btn btn-primary btn-block">Simpan & Update Stok</button>
        </form>
    `);
    document.getElementById('p-item').addEventListener('change', function () {
        document.getElementById('p-newname-group').classList.toggle('hidden', this.value !== '__new__');
    });
}
let pendingInvoiceData = null;
function previewInvoice(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (ev) {
        pendingInvoiceData = ev.target.result;
        const img = document.getElementById('p-invoice-preview');
        img.src = pendingInvoiceData;
        img.classList.remove('hidden');
    };
    reader.readAsDataURL(file);
}
function submitPurchaseForm(e) {
    e.preventDefault();
    const db = getDB();
    const itemSel = document.getElementById('p-item').value;
    const qty = Number(document.getElementById('p-qty').value);
    const price = Number(document.getElementById('p-price').value);
    let item;
    if (itemSel === '__new__') {
        const name = document.getElementById('p-newname').value.trim();
        if (!name) { toast('Nama barang baru wajib diisi.', 'danger'); return; }
        item = { id: uid('itm'), name, category: '', stock: 0, unit: 'pcs', location: '', rack: '', minStock: 1 };
        db.items.push(item);
    } else {
        item = db.items.find(i => i.id === itemSel);
    }
    item.stock += qty;
    db.purchases.push({
        id: uid('pur'), itemId: item.id, itemName: item.name, qty, price,
        supplier: document.getElementById('p-supplier').value.trim(),
        date: document.getElementById('p-date').value,
        invoicePhoto: pendingInvoiceData || null,
        addedBy: currentUser().name
    });
    pendingInvoiceData = null;
    saveDB(db);
    closeModal();
    toast('Pembelian dicatat & stok barang diperbarui.');
    renderPage('purchases');
}
function viewInvoice(purchaseId) {
    const db = getDB();
    const p = db.purchases.find(x => x.id === purchaseId);
    if (!p || !p.invoicePhoto) return;
    openModal('Foto Invoice - ' + p.itemName, `<img src="${p.invoicePhoto}" style="width:100%;border-radius:8px;">`);
}

/* =========================================================
   PENGAJUAN PEMBELIAN (bisa dicetak / print untuk atasan)
   ========================================================= */
function renderRequestsPage() {
    const db = getDB();
    const isManagerLevel = hasRole('admin', 'manajemen');
    const list = isManagerLevel ? db.requests : db.requests.filter(r => r.requestedBy === currentUser().name);
    const rows = [...list].sort((a, b) => new Date(b.date) - new Date(a.date));
    return `
    <div class="page-header">
        <p>Ajukan permintaan pembelian barang. ${hasRole('manajemen') ? 'Sebagai Manajemen, Anda dapat menyetujui/menolak pengajuan.' : ''}</p>
        <button class="btn btn-primary" onclick="openRequestForm()"><i data-lucide="plus"></i>Buat Pengajuan</button>
    </div>
    <div class="table-container">
        <table class="data-table">
            <thead><tr><th>Tanggal</th><th>Barang</th><th>Jml</th><th>Diajukan Oleh</th><th>Status</th><th>Aksi</th></tr></thead>
            <tbody>
                ${rows.length ? rows.map(r => `<tr>
                    <td>${fmtDate(r.date)}</td><td>${escapeHtml(r.itemName)}</td><td>${r.qty}</td>
                    <td>${escapeHtml(r.requestedBy)}</td><td>${requestStatusBadge(r.status)}</td>
                    <td class="actions">
                        ${hasRole('manajemen') && r.status === 'pending' ? `
                            <button class="btn-icon" title="Setujui" onclick="updateRequestStatus('${r.id}','disetujui')"><i data-lucide="check"></i></button>
                            <button class="btn-icon text-danger" title="Tolak" onclick="updateRequestStatus('${r.id}','ditolak')"><i data-lucide="x"></i></button>
                        ` : ''}
                        <button class="btn-icon" title="Cetak Form Pengajuan" onclick="printRequest('${r.id}')"><i data-lucide="printer"></i></button>
                    </td>
                </tr>`).join('') : `<tr class="empty-row"><td colspan="6">Belum ada pengajuan.</td></tr>`}
            </tbody>
        </table>
    </div>`;
}
function requestStatusBadge(status) {
    if (status === 'disetujui') return '<span class="badge badge-success">Disetujui</span>';
    if (status === 'ditolak') return '<span class="badge badge-danger">Ditolak</span>';
    return '<span class="badge badge-warning">Menunggu</span>';
}
function openRequestForm() {
    openModal('Buat Pengajuan Pembelian', `
        <form id="request-form" onsubmit="submitRequestForm(event)">
            <div class="form-group"><label>Nama Barang</label><input type="text" id="r-name" required></div>
            <div class="form-row">
                <div class="form-group"><label>Jumlah</label><input type="number" id="r-qty" min="1" value="1" required></div>
                <div class="form-group"><label>Perkiraan Harga (Rp, opsional)</label><input type="number" id="r-price" min="0" value="0"></div>
            </div>
            <div class="form-group"><label>Alasan / Keterangan</label><textarea id="r-reason" placeholder="cth: stok habis, kebutuhan proyek X" required></textarea></div>
            <button type="submit" class="btn btn-primary btn-block">Ajukan</button>
        </form>
    `);
}
function submitRequestForm(e) {
    e.preventDefault();
    const db = getDB();
    db.requests.push({
        id: uid('req'), itemName: document.getElementById('r-name').value.trim(),
        qty: Number(document.getElementById('r-qty').value),
        estPrice: Number(document.getElementById('r-price').value) || 0,
        reason: document.getElementById('r-reason').value.trim(),
        requestedBy: currentUser().name, date: todayStr(), status: 'pending'
    });
    saveDB(db);
    closeModal();
    toast('Pengajuan pembelian berhasil dikirim.');
    renderPage('requests');
    renderSidebar();
}
function updateRequestStatus(reqId, status) {
    const db = getDB();
    const r = db.requests.find(x => x.id === reqId);
    r.status = status;
    r.approvedBy = currentUser().name;
    r.decisionDate = todayStr();
    saveDB(db);
    toast(status === 'disetujui' ? 'Pengajuan disetujui.' : 'Pengajuan ditolak.', status === 'disetujui' ? 'success' : 'warning');
    renderPage('requests');
    renderSidebar();
}
// Cetak form pengajuan pembelian untuk diajukan ke level lebih atas
function printRequest(reqId) {
    const db = getDB();
    const r = db.requests.find(x => x.id === reqId);
    if (!r) return;
    const html = `
        <div style="text-align:center;margin-bottom:24px;">
            <h2 style="margin-bottom:4px;">FORM PENGAJUAN PEMBELIAN BARANG</h2>
            <p style="color:#555;">GudangKita - Sistem Manajemen Stok & Peminjaman</p>
        </div>
        <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
            <tr><td style="padding:6px 0;width:200px;"><strong>No. Pengajuan</strong></td><td>: ${escapeHtml(r.id)}</td></tr>
            <tr><td style="padding:6px 0;"><strong>Tanggal Pengajuan</strong></td><td>: ${fmtDate(r.date)}</td></tr>
            <tr><td style="padding:6px 0;"><strong>Diajukan Oleh</strong></td><td>: ${escapeHtml(r.requestedBy)}</td></tr>
            <tr><td style="padding:6px 0;"><strong>Nama Barang</strong></td><td>: ${escapeHtml(r.itemName)}</td></tr>
            <tr><td style="padding:6px 0;"><strong>Jumlah</strong></td><td>: ${r.qty}</td></tr>
            <tr><td style="padding:6px 0;"><strong>Perkiraan Harga</strong></td><td>: ${r.estPrice ? fmtRupiah(r.estPrice) : '-'}</td></tr>
            <tr><td style="padding:6px 0;vertical-align:top;"><strong>Alasan / Keterangan</strong></td><td>: ${escapeHtml(r.reason)}</td></tr>
            <tr><td style="padding:6px 0;"><strong>Status</strong></td><td>: ${r.status === 'disetujui' ? 'DISETUJUI' : r.status === 'ditolak' ? 'DITOLAK' : 'MENUNGGU PERSETUJUAN'}</td></tr>
            ${r.approvedBy ? `<tr><td style="padding:6px 0;"><strong>Diputuskan Oleh</strong></td><td>: ${escapeHtml(r.approvedBy)} (${fmtDate(r.decisionDate)})</td></tr>` : ''}
        </table>
        <table style="width:100%;margin-top:60px;border-collapse:collapse;">
            <tr>
                <td style="width:33%;text-align:center;">Diajukan oleh,<br><br><br><br>(...........................)</td>
                <td style="width:33%;text-align:center;">Diperiksa Admin,<br><br><br><br>(...........................)</td>
                <td style="width:33%;text-align:center;">Disetujui Manajemen,<br><br><br><br>(...........................)</td>
            </tr>
        </table>
    `;
    document.getElementById('print-area').innerHTML = html;
    window.print();
}

/* =========================================================
   BARANG RUSAK / MT
   ========================================================= */
function renderDamagesPage() {
    if (!hasRole('admin', 'manajemen')) return `<p class="text-muted">Anda tidak memiliki akses ke halaman ini.</p>`;
    const db = getDB();
    const rows = [...db.damages].sort((a, b) => new Date(b.date) - new Date(a.date));
    return `
    <div class="page-header">
        <p>Catatan barang rusak / perlu maintenance, baik dari pengembalian pinjaman maupun laporan langsung.</p>
        <button class="btn btn-primary" onclick="openDamageForm()"><i data-lucide="plus"></i>Catat Barang Rusak/MT</button>
    </div>
    <div class="table-container">
        <table class="data-table">
            <thead><tr><th>Tanggal</th><th>Barang</th><th>Jml</th><th>Jenis</th><th>Deskripsi</th><th>Sumber</th><th>Status</th><th>Aksi</th></tr></thead>
            <tbody>
                ${rows.length ? rows.map(d => `<tr>
                    <td>${fmtDate(d.date)}</td><td>${escapeHtml(d.itemName)}</td><td>${d.qty}</td>
                    <td>${d.type === 'mt' ? '<span class="badge badge-warning">MT</span>' : '<span class="badge badge-danger">Rusak</span>'}</td>
                    <td>${escapeHtml(d.description || '-')}</td>
                    <td>${d.source === 'pengembalian' ? 'Pengembalian' : 'Laporan Langsung'}</td>
                    <td>${damageStatusBadge(d.status)}</td>
                    <td>
                        ${d.status !== 'selesai' ? `<select onchange="updateDamageStatus('${d.id}', this.value)" class="btn-sm">
                            <option value="">Ubah status...</option>
                            <option value="proses">Sedang Diproses</option>
                            <option value="selesai">Selesai Diperbaiki</option>
                        </select>` : '-'}
                    </td>
                </tr>`).join('') : `<tr class="empty-row"><td colspan="8">Belum ada catatan barang rusak/MT.</td></tr>`}
            </tbody>
        </table>
    </div>`;
}
function damageStatusBadge(s) {
    if (s === 'selesai') return '<span class="badge badge-success">Selesai</span>';
    if (s === 'proses') return '<span class="badge badge-info">Diproses</span>';
    return '<span class="badge badge-warning">Menunggu</span>';
}
function openDamageForm() {
    const db = getDB();
    const options = db.items.map(i => `<option value="${i.id}">${escapeHtml(i.name)}</option>`).join('');
    openModal('Catat Barang Rusak / MT', `
        <form id="damage-form" onsubmit="submitDamageForm(event)">
            <div class="form-group"><label>Barang</label><select id="d-item" required>${options}</select></div>
            <div class="form-row">
                <div class="form-group"><label>Jumlah</label><input type="number" id="d-qty" min="1" value="1" required></div>
                <div class="form-group"><label>Jenis</label>
                    <select id="d-type" required><option value="rusak">Rusak</option><option value="mt">Perlu Maintenance (MT)</option></select>
                </div>
            </div>
            <div class="form-group"><label>Deskripsi</label><textarea id="d-desc" required></textarea></div>
            <button type="submit" class="btn btn-primary btn-block">Simpan Catatan</button>
        </form>
    `);
}
function submitDamageForm(e) {
    e.preventDefault();
    const db = getDB();
    const item = db.items.find(i => i.id === document.getElementById('d-item').value);
    const qty = Number(document.getElementById('d-qty').value);
    if (item && item.stock >= qty) item.stock -= qty; // pindahkan dari stok baik ke catatan rusak/MT
    db.damages.push({
        id: uid('dmg'), itemId: item.id, itemName: item.name, qty,
        type: document.getElementById('d-type').value,
        description: document.getElementById('d-desc').value.trim(),
        source: 'laporan langsung', reportedBy: currentUser().name, date: todayStr(), status: 'menunggu'
    });
    saveDB(db);
    closeModal();
    toast('Catatan barang rusak/MT disimpan.', 'warning');
    renderPage('damages');
}
function updateDamageStatus(id, status) {
    if (!status) return;
    const db = getDB();
    const d = db.damages.find(x => x.id === id);
    d.status = status;
    saveDB(db);
    renderPage('damages');
    toast('Status diperbarui.');
}

/* =========================================================
   BARANG HILANG (identifikasi)
   ========================================================= */
function renderLostPage() {
    if (!hasRole('admin', 'manajemen')) return `<p class="text-muted">Anda tidak memiliki akses ke halaman ini.</p>`;
    const db = getDB();
    const rows = [...db.lost].sort((a, b) => new Date(b.date) - new Date(a.date));
    return `
    <div class="page-header">
        <p>Identifikasi dan catat barang yang hilang dari gudang.</p>
        <button class="btn btn-primary" onclick="openLostForm()"><i data-lucide="plus"></i>Laporkan Barang Hilang</button>
    </div>
    <div class="table-container">
        <table class="data-table">
            <thead><tr><th>Tanggal</th><th>Barang</th><th>Jml</th><th>Lokasi Terakhir</th><th>Keterangan Identifikasi</th><th>Dilaporkan Oleh</th><th>Status</th></tr></thead>
            <tbody>
                ${rows.length ? rows.map(l => `<tr>
                    <td>${fmtDate(l.date)}</td><td>${escapeHtml(l.itemName)}</td><td>${l.qty}</td>
                    <td>${escapeHtml(l.lastLocation || '-')}</td><td>${escapeHtml(l.description || '-')}</td>
                    <td>${escapeHtml(l.reportedBy)}</td>
                    <td>
                        <select onchange="updateLostStatus('${l.id}', this.value)" class="btn-sm">
                            <option value="dicari" ${l.status === 'dicari' ? 'selected' : ''}>Sedang Dicari</option>
                            <option value="ditemukan" ${l.status === 'ditemukan' ? 'selected' : ''}>Ditemukan</option>
                            <option value="writeoff" ${l.status === 'writeoff' ? 'selected' : ''}>Write-off (Hilang Permanen)</option>
                        </select>
                    </td>
                </tr>`).join('') : `<tr class="empty-row"><td colspan="7">Belum ada laporan barang hilang.</td></tr>`}
            </tbody>
        </table>
    </div>`;
}
function openLostForm() {
    const db = getDB();
    const options = db.items.map(i => `<option value="${i.id}">${escapeHtml(i.name)}</option>`).join('');
    openModal('Laporkan Barang Hilang', `
        <form id="lost-form" onsubmit="submitLostForm(event)">
            <div class="form-group"><label>Barang</label><select id="l-item" required>${options}</select></div>
            <div class="form-row">
                <div class="form-group"><label>Jumlah</label><input type="number" id="l-qty" min="1" value="1" required></div>
                <div class="form-group"><label>Lokasi Terakhir Terlihat</label><input type="text" id="l-location"></div>
            </div>
            <div class="form-group"><label>Keterangan Identifikasi</label><textarea id="l-desc" placeholder="ciri-ciri barang, dugaan penyebab hilang, dsb." required></textarea></div>
            <button type="submit" class="btn btn-primary btn-block">Simpan Laporan</button>
        </form>
    `);
}
function submitLostForm(e) {
    e.preventDefault();
    const db = getDB();
    const item = db.items.find(i => i.id === document.getElementById('l-item').value);
    const qty = Number(document.getElementById('l-qty').value);
    if (item && item.stock >= qty) item.stock -= qty;
    db.lost.push({
        id: uid('lost'), itemId: item.id, itemName: item.name, qty,
        lastLocation: document.getElementById('l-location').value.trim(),
        description: document.getElementById('l-desc').value.trim(),
        reportedBy: currentUser().name, date: todayStr(), status: 'dicari'
    });
    saveDB(db);
    closeModal();
    toast('Laporan barang hilang disimpan.', 'warning');
    renderPage('lost');
}
function updateLostStatus(id, status) {
    const db = getDB();
    const l = db.lost.find(x => x.id === id);
    l.status = status;
    saveDB(db);
    toast('Status barang hilang diperbarui.');
    renderPage('lost');
}

/* =========================================================
   MANAJEMEN USER (khusus role manajemen)
   ========================================================= */
function renderUsersPage() {
    if (!hasRole('manajemen')) return `<p class="text-muted">Anda tidak memiliki akses ke halaman ini.</p>`;
    const db = getDB();
    return `
    <div class="page-header">
        <p>Kelola akun pengguna sistem (staff, admin, manajemen).</p>
        <button class="btn btn-primary" onclick="openUserForm()"><i data-lucide="plus"></i>Tambah User</button>
    </div>
    <div class="table-container">
        <table class="data-table">
            <thead><tr><th>Nama</th><th>Username</th><th>Role</th><th>Aksi</th></tr></thead>
            <tbody>
                ${db.users.map(u => `<tr>
                    <td>${escapeHtml(u.name)}</td><td>${escapeHtml(u.username)}</td>
                    <td>${roleBadge(u.role)}</td>
                    <td class="actions">
                        <button class="btn-icon" title="Edit" onclick="openUserForm('${u.id}')"><i data-lucide="pencil"></i></button>
                        ${u.id !== currentUser().id ? `<button class="btn-icon text-danger" title="Hapus" onclick="deleteUser('${u.id}')"><i data-lucide="trash-2"></i></button>` : ''}
                    </td>
                </tr>`).join('')}
            </tbody>
        </table>
    </div>`;
}
function roleBadge(role) {
    if (role === 'manajemen') return '<span class="badge badge-purple">Manajemen</span>';
    if (role === 'admin') return '<span class="badge badge-info">Admin</span>';
    return '<span class="badge badge-muted">Staff</span>';
}
function openUserForm(userId) {
    const db = getDB();
    const u = userId ? db.users.find(x => x.id === userId) : null;
    openModal(u ? 'Edit User' : 'Tambah User', `
        <form id="user-form" onsubmit="submitUserForm(event, '${userId || ''}')">
            <div class="form-group"><label>Nama Lengkap</label><input type="text" id="u-name" value="${u ? escapeHtml(u.name) : ''}" required></div>
            <div class="form-group"><label>Username</label><input type="text" id="u-username" value="${u ? escapeHtml(u.username) : ''}" required></div>
            <div class="form-group"><label>Password ${u ? '(kosongkan jika tidak diubah)' : ''}</label><input type="text" id="u-password" ${u ? '' : 'required'}></div>
            <div class="form-group"><label>Role</label>
                <select id="u-role" required>
                    <option value="staff" ${u && u.role === 'staff' ? 'selected' : ''}>Staff</option>
                    <option value="admin" ${u && u.role === 'admin' ? 'selected' : ''}>Admin</option>
                    <option value="manajemen" ${u && u.role === 'manajemen' ? 'selected' : ''}>Manajemen</option>
                </select>
            </div>
            <button type="submit" class="btn btn-primary btn-block">Simpan</button>
        </form>
    `);
}
function submitUserForm(e, userId) {
    e.preventDefault();
    const db = getDB();
    const username = document.getElementById('u-username').value.trim();
    const dupe = db.users.find(u => u.username === username && u.id !== userId);
    if (dupe) { toast('Username sudah digunakan.', 'danger'); return; }
    const pass = document.getElementById('u-password').value;
    if (userId) {
        const idx = db.users.findIndex(u => u.id === userId);
        db.users[idx].name = document.getElementById('u-name').value.trim();
        db.users[idx].username = username;
        db.users[idx].role = document.getElementById('u-role').value;
        if (pass) db.users[idx].password = pass;
        toast('User berhasil diperbarui.');
    } else {
        db.users.push({
            id: uid('u'), name: document.getElementById('u-name').value.trim(),
            username, password: pass, role: document.getElementById('u-role').value
        });
        toast('User baru berhasil ditambahkan.');
    }
    saveDB(db);
    closeModal();
    renderPage('users');
}
function deleteUser(userId) {
    if (!confirm('Hapus user ini?')) return;
    const db = getDB();
    db.users = db.users.filter(u => u.id !== userId);
    saveDB(db);
    renderPage('users');
    toast('User dihapus.', 'warning');
}

/* =========================================================
   AUTH & INISIALISASI
   ========================================================= */
function doLogin(e) {
    e.preventDefault();
    const username = document.getElementById('login-username').value.trim();
    const password = document.getElementById('login-password').value;
    const db = getDB();
    const user = db.users.find(u => u.username === username && u.password === password);
    if (!user) { toast('Username atau password salah.', 'danger'); return; }
    setSession(user);
    enterApp();
}
function doLogout() {
    clearSession();
    document.getElementById('app-layout').classList.add('hidden');
    document.getElementById('login-screen').classList.remove('hidden');
    document.getElementById('login-form').reset();
}
function enterApp() {
    const u = currentUser();
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('app-layout').classList.remove('hidden');
    document.getElementById('user-display-name').textContent = u.name;
    document.getElementById('user-display-role').textContent = u.role;
    currentPage = 'dashboard';
    renderSidebar();
    navigateTo('dashboard');
}
function updateClock() {
    const el = document.getElementById('live-clock');
    if (el) el.textContent = new Date().toLocaleString('id-ID', { weekday: 'short', day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', second: '2-digit' });
}
function initTheme() {
    const saved = localStorage.getItem('gudangkita_theme');
    if (saved === 'dark') {
        document.body.classList.remove('light-mode'); document.body.classList.add('dark-mode');
        document.querySelector('#theme-toggle i').setAttribute('data-lucide', 'sun');
    }
}
function toggleTheme() {
    const isDark = document.body.classList.toggle('dark-mode');
    document.body.classList.toggle('light-mode', !isDark);
    localStorage.setItem('gudangkita_theme', isDark ? 'dark' : 'light');
    document.querySelector('#theme-toggle i').setAttribute('data-lucide', isDark ? 'sun' : 'moon');
    if (window.lucide) lucide.createIcons();
}

document.addEventListener('DOMContentLoaded', () => {
    if (window.lucide) lucide.createIcons();
    initTheme();

    document.getElementById('login-form').addEventListener('submit', doLogin);
    document.getElementById('logout-btn').addEventListener('click', doLogout);
    document.getElementById('theme-toggle').addEventListener('click', toggleTheme);
    document.getElementById('modal-close').addEventListener('click', closeModal);
    document.getElementById('modal').addEventListener('click', (e) => { if (e.target.id === 'modal') closeModal(); });
    document.getElementById('mobile-toggle-open').addEventListener('click', () => document.getElementById('sidebar').classList.add('open'));
    document.getElementById('mobile-toggle-close').addEventListener('click', () => document.getElementById('sidebar').classList.remove('open'));

    setInterval(updateClock, 1000);
    updateClock();

    const session = getSession();
    if (session) enterApp();
});
