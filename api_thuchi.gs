/* ============================================================================
 * API BÁO CÁO THU CHI - PHI LONG AUTO
 * ----------------------------------------------------------------------------
 * Nguồn: file "BÁO CÁO NGÀY" (bản LIVE nhân bản từ app) - tab THU_CHI
 * Dùng cho trang: trolyoto.github.io/appsheet/philongauto.baocaothuchi.html
 *
 * Lý do có API này (25/10/2026): trang cũ nhận dữ liệu qua tham số URL do
 * AppSheet ghép (&tbh=&tk=&cncc=&ck=). Link thực tế dài ~34.000 ký tự, vượt
 * giới hạn ~8.150 ký tự của GitHub Pages -> lỗi HTTP 414. Ngoài ra chuỗi app
 * gửi chỉ có 5 cột (ngày / THU-CHI / số phiếu / số tiền / hình thức) nên bảng
 * bị lệch cột và không có tên khách hàng, diễn giải, phân nhóm.
 *
 * GỌI:
 *   ?tungay=18/10/2026&denngay=24/10/2026&chinhanh=CN1
 *   ?ngay=18/10/26 - 24/10/26&chinhanh=CN1 , CN2 , CN3     (đúng dạng AppSheet đang gửi)
 *   ?tungay=...&denngay=...                                 (trống chi nhánh = tất cả)
 *
 * TRẢ:
 *   { ok:true, ky:{tu,den}, chinhanh, nguon, tinhLuc, tongDong, soDong,
 *     boQuaKhongNgay, rows:[ {ngay,loai,sophieu,nhom,doitac,tien,lydo,ht,cn,iddh} ] }
 * ==========================================================================*/

var FILE_ID = '1N1ngzYfXP2vIr_56cMGp_5WyZfmAH2NQQtm1jDkEhks';   // BÁO CÁO NGÀY
var TAB      = 'THU_CHI';

function doGet(e) {
  try {
    return json(apiThuChi((e && e.parameter) || {}));
  } catch (err) {
    return json({ ok: false, loi: String(err && err.message ? err.message : err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────────── SỐ & NGÀY ─────────────────────────────── */
/* Chịu được: 50.000 | 50,000 | 50000 | "-50000" | (50.000) | số thật */
function parseNum(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number') return raw;
  var t = String(raw).trim();
  if (!t || t === '-' || t === '—') return 0;
  var am = /^\(.*\)$/.test(t) || /^-/.test(t);
  t = t.replace(/[()\-\s]/g, '').replace(/[^0-9.,]/g, '');
  if (!t) return 0;
  var coC = t.indexOf(',') > -1, coD = t.indexOf('.') > -1, v;
  if (coC && coD) {
    if (t.lastIndexOf(',') > t.lastIndexOf('.')) v = parseFloat(t.replace(/\./g, '').replace(',', '.'));
    else v = parseFloat(t.replace(/,/g, ''));
  } else if (coC || coD) {
    var d = coC ? ',' : '.';
    var k = t.split(d);
    /* nhiều dấu, hoặc 1 dấu + đúng 3 chữ số phía sau => dấu ngăn hàng nghìn */
    if (k.length > 2 || k[k.length - 1].length === 3) v = parseFloat(k.join(''));
    else v = parseFloat(k.join('.'));
  } else v = parseFloat(t);
  return isFinite(v) ? (am ? -v : v) : 0;
}

function p2(n) { return (n < 10 ? '0' : '') + n; }
function d2s(d) { return p2(d.getDate()) + '/' + p2(d.getMonth() + 1) + '/' + d.getFullYear(); }

/* Nhận Date thật, "dd/MM/yyyy", "dd/MM/yy", "yyyy-MM-dd" */
function toDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) { var y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1]); }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

/* ───────────────────────────── API CHÍNH ─────────────────────────────── */
function apiThuChi(p) {
  /* ── 1. Kỳ báo cáo ── */
  var tu = toDate(p.tungay), den = toDate(p.denngay);
  var ngay = String(p.ngay || '').trim();
  if ((!tu || !den) && ngay) {
    /* "18/10/26 - 24/10/26"  hoặc  "18/10/2026" */
    var ph = ngay.split(/\s+-\s+|\s-\s/);
    if (ph.length >= 2) { tu = tu || toDate(ph[0]); den = den || toDate(ph[1]); }
    else { tu = tu || toDate(ngay); den = den || toDate(ngay); }
  }
  var hnay = new Date();
  if (!den) den = new Date(hnay.getFullYear(), hnay.getMonth(), hnay.getDate());
  if (!tu) tu = den;
  if (tu.getTime() > den.getTime()) { var x = tu; tu = den; den = x; }

  /* ── 2. Chi nhánh: "CN1 , CN2 , CN3" / "CN1" / trống = tất cả ── */
  var cnRaw = String(p.chinhanh || '').toUpperCase();
  var cnList = cnRaw.split(/[,;|]+/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
  var locCN = {};
  var coLocCN = cnList.length > 0;
  cnList.forEach(function (c) { locCN[c] = 1; });

  /* ── 3. Đọc tab THU_CHI ── */
  var sh = SpreadsheetApp.openById(FILE_ID).getSheetByName(TAB);
  if (!sh) return { ok: false, loi: 'Không tìm thấy tab ' + TAB + ' trong file BÁO CÁO NGÀY.' };
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { ok: false, loi: 'Tab ' + TAB + ' không có dữ liệu.' };

  var hdr = vals[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  function idx(name) { return hdr.indexOf(name); }
  var iLoai = idx('loai_phieu'), iSo = idx('id_tc'), iNhom = idx('phan_nhom'),
      iDT = idx('doi_tac'), iKH = idx('ten_kh'), iNCC = idx('ncc'),
      iTien = idx('thanh_tien'), iLyDo = idx('ly_do'), iHT = idx('hinh_thuc_tt'),
      iNgay = idx('ngay_tao'), iCN = idx('chi_nhanh'), iDH = idx('id_dh'),
      iNguoi = idx('nguoi_tao');
  if (iLoai < 0 || iTien < 0 || iNgay < 0) {
    return { ok: false, loi: 'Tab THU_CHI thiếu cột bắt buộc (loai_phieu / thanh_tien / ngay_tao).' };
  }

  var rows = [], boQua = 0, tongDong = vals.length - 1;
  var tTu = tu.getTime(), tDen = den.getTime();

  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var loai = String(row[iLoai] || '').trim().toUpperCase();
    if (loai !== 'THU' && loai !== 'CHI') continue;

    var d = toDate(row[iNgay]);
    if (!d) { boQua++; continue; }
    var t = d.getTime();
    if (t < tTu || t > tDen) continue;

    var cn = String(row[iCN] || '').trim().toUpperCase();
    if (coLocCN && !locCN[cn]) continue;

    var ten = iDT >= 0 ? String(row[iDT] || '').trim() : '';
    if (!ten && iKH >= 0) ten = String(row[iKH] || '').trim();
    if (!ten && iNCC >= 0) ten = String(row[iNCC] || '').trim();

    rows.push({
      ngay: d2s(d),
      ts: t,
      loai: loai,
      sophieu: iSo >= 0 ? String(row[iSo] || '').trim() : '',
      nhom: iNhom >= 0 ? String(row[iNhom] || '').trim() : '',
      doitac: ten,
      ncc: iNCC >= 0 ? String(row[iNCC] || '').trim() : '',
      tien: parseNum(row[iTien]),
      lydo: iLyDo >= 0 ? String(row[iLyDo] || '').trim() : '',
      ht: (iHT >= 0 ? String(row[iHT] || '').trim().toUpperCase() : '') || 'TM',
      cn: cn,
      iddh: iDH >= 0 ? String(row[iDH] || '').trim() : '',
      nguoi: iNguoi >= 0 ? String(row[iNguoi] || '').trim() : ''
    });
  }

  rows.sort(function (a, b) { return a.ts - b.ts || (a.sophieu < b.sophieu ? -1 : 1); });

  return {
    ok: true,
    ky: { tu: d2s(tu), den: d2s(den) },
    chinhanh: coLocCN ? cnList.join(', ') : '',
    nguon: 'BÁO CÁO NGÀY / tab THU_CHI',
    tinhLuc: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm'),
    tongDong: tongDong,
    soDong: rows.length,
    boQuaKhongNgay: boQua,
    rows: rows
  };
}

/* Chạy tay để kiểm tra nhanh trong editor */
function test_() {
  var kq = apiThuChi({ ngay: '18/10/26 - 24/10/26', chinhanh: 'CN1 , CN2 , CN3' });
  var thu = 0, chi = 0;
  kq.rows.forEach(function (x) { if (x.loai === 'THU') thu += x.tien; else chi += x.tien; });
  Logger.log('%s dòng | THU %s | CHI %s | bỏ qua %s',
    kq.soDong, thu.toLocaleString('vi-VN'), chi.toLocaleString('vi-VN'), kq.boQuaKhongNgay);
}
