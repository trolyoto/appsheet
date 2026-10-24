/**
 * PHI LONG AUTO — API CONG NO PHAI THU KHACH HANG (TK 131)
 * Nguon: Google Sheet "BAO CAO NGAY"
 *   - tab "SO CONG NO KH"            : so chi tiet tung chung tu (nguon tinh toan)
 *   - tab "BC TONG NO 3 CHI NHANH"   : ma KH, chi nhanh quan ly, no hien tai tren APP
 *
 * Cong thuc chuan TK 131:
 *   No dau ky  = tong (Phat sinh - Da thu) cua cac chung tu TRUOC "tu ngay"
 *   Phat sinh  = tong Phat sinh trong [tu ngay .. den ngay]
 *   Da thu     = tong Da thu    trong [tu ngay .. den ngay]
 *   Cuoi ky    = Dau ky + Phat sinh - Da thu
 *
 * Goi:
 *   .../exec                              -> lay thang BC TONG NO (so chuan, ky theo file)
 *   .../exec?tungay=01/03/2026&denngay=30/09/2026   -> tinh lai tu SO CONG NO KH
 *   .../exec?...&chinhanh=CN1
 *   .../exec?debug=1                      -> in ra ten tab + dong tieu de de tra soat
 *   .../exec?...&ct=1                     -> kem bang chi tiet chung tu
 */

var SS_ID      = '1N1ngzYfXP2vIr_56cMGp_5WyZfmAH2NQQtm1jDkEhks';
var TAB_SO     = 'SỔ CÔNG NỢ KH';
var TAB_TH     = 'BC TỔNG NỢ 3 CHI NHÁNH';
var CACHE_GIAY = 60;

/* ───────────────────────── tien ich ───────────────────────── */

function khongDau_(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function soVN_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  var t = String(v).trim();
  if (!t || t === '-' || t === '—') return 0;
  var am = /^\(.*\)$/.test(t) || /^-/.test(t);
  t = t.replace(/[()\-\s]/g, '').replace(/[^0-9.,]/g, '');
  if (!t) return 0;
  var coC = t.indexOf(',') > -1, coD = t.indexOf('.') > -1, v2;
  if (coC && coD) {
    if (t.lastIndexOf(',') > t.lastIndexOf('.')) v2 = parseFloat(t.replace(/\./g, '').replace(',', '.'));
    else v2 = parseFloat(t.replace(/,/g, ''));
  } else if (coC || coD) {
    var d = coC ? ',' : '.';
    var k = t.split(d);
    if (k.length > 2 || k[k.length - 1].length === 3) v2 = parseFloat(k.join(''));
    else v2 = parseFloat(k.join('.'));
  } else v2 = parseFloat(t);
  if (!isFinite(v2)) return 0;
  return am ? -v2 : v2;
}

/* dd/MM/yyyy | dd/MM/yy | yyyy-MM-dd | Date -> moc 00:00 theo ngay */
function docNgay_(v) {
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    var d = +m[1], th = +m[2], n = +m[3];
    if (n < 100) n += 2000;
    return new Date(n, th - 1, d);
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var dd = new Date(s);
  return isNaN(dd.getTime()) ? null : new Date(dd.getFullYear(), dd.getMonth(), dd.getDate());
}

function inNgay_(d) {
  if (!d) return '';
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear();
}

function layTab_(ss, ten) {
  var sh = ss.getSheetByName(ten);
  if (sh) return sh;
  var dich = khongDau_(ten), ds = ss.getSheets(), i;
  for (i = 0; i < ds.length; i++) if (khongDau_(ds[i].getName()) === dich) return ds[i];
  for (i = 0; i < ds.length; i++) if (khongDau_(ds[i].getName()).indexOf(dich) > -1) return ds[i];
  return null;
}

/* tim chi so cot theo ten tieu de (khong dau, khong phan biet hoa thuong) */
function timCot_(hang, ds) {
  var i, j, h;
  for (i = 0; i < hang.length; i++) {
    h = khongDau_(hang[i]);
    if (!h) continue;
    for (j = 0; j < ds.length; j++) if (h === khongDau_(ds[j])) return i;
  }
  for (i = 0; i < hang.length; i++) {
    h = khongDau_(hang[i]);
    if (!h) continue;
    for (j = 0; j < ds.length; j++) if (h.indexOf(khongDau_(ds[j])) > -1) return i;
  }
  return -1;
}

/* dong tieu de = dong co nhieu o chu nhat trong 8 dong dau */
function timHangTieuDe_(bang, tuKhoa) {
  var tot = 0, diem = -1, i, j, d;
  for (i = 0; i < Math.min(8, bang.length); i++) {
    d = 0;
    for (j = 0; j < bang[i].length; j++) {
      var o = khongDau_(bang[i][j]);
      if (!o) continue;
      for (var k = 0; k < tuKhoa.length; k++) if (o.indexOf(khongDau_(tuKhoa[k])) > -1) { d++; break; }
    }
    if (d > diem) { diem = d; tot = i; }
  }
  return diem > 0 ? tot : 0;
}

function traJSON_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── diem vao ───────────────────────── */

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  try {
    if (p.debug) return traJSON_(chanDoan_());

    var ss = SpreadsheetApp.openById(SS_ID);
    var kyFile = docKyFile_(ss);            /* ky dang dat trong file BAO CAO NGAY */
    var kemCT  = !!p.ct;

    var tu = docNgay_(p.tungay), den = docNgay_(p.denngay);
    /* khong chon ky, hoac chon trung ky cua file -> lay thang bang chuan BC TONG NO */
    var dungBC = (!tu || !den) ||
                 (kyFile.tu && kyFile.den &&
                  tu.getTime() === kyFile.tu.getTime() &&
                  den.getTime() === kyFile.den.getTime());

    if (!tu) tu = kyFile.tu;
    if (!den) den = kyFile.den;
    if (!den) den = new Date();
    if (!tu) tu = new Date(den.getFullYear(), den.getMonth(), 1);
    if (tu.getTime() > den.getTime()) { var tmp = tu; tu = den; den = tmp; }

    var cn = khongDau_(p.chinhanh || '');
    if (cn === 'ALL' || cn === 'TAT CA' || cn === 'TATCA') cn = '';

    var khoa = 'cnkh_v2|' + (dungBC ? 'BC' : 'SO') + '|' + inNgay_(tu) + '|' + inNgay_(den) + '|' + cn + '|' + (kemCT ? 1 : 0);
    var cache = CacheService.getScriptCache();
    var san = cache.get(khoa);
    if (san) return ContentService.createTextOutput(san).setMimeType(ContentService.MimeType.JSON);

    var kq = dungBC ? docBangChuan_(ss, tu, den, cn, kemCT)
                    : tinhCongNo_(tu, den, cn, kemCT);
    kq.kyFile = { tu: inNgay_(kyFile.tu), den: inNgay_(kyFile.den) };
    var chuoi = JSON.stringify(kq);
    if (chuoi.length < 95000) { try { cache.put(khoa, chuoi, CACHE_GIAY); } catch (x) {} }
    return ContentService.createTextOutput(chuoi).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return traJSON_({ ok: false, loi: String(err && err.message ? err.message : err) });
  }
}

/* ky dang dat trong tab BC TONG NO (o "TU NGAY:" / "DEN NGAY:") */
function docKyFile_(ss) {
  var ra = { tu: null, den: null };
  var sh = layTab_(ss, TAB_TH);
  if (!sh) return ra;
  var r = Math.min(sh.getLastRow(), 6), c = Math.min(sh.getLastColumn(), 8);
  if (r < 1) return ra;
  var bang = sh.getRange(1, 1, r, c).getDisplayValues();
  for (var i = 0; i < bang.length; i++) {
    for (var j = 0; j < bang[i].length; j++) {
      var o = khongDau_(bang[i][j]);
      if (o.indexOf('TU NGAY') === 0 || o.indexOf('DEN NGAY') === 0) {
        for (var k = j + 1; k < bang[i].length; k++) {
          var d = docNgay_(bang[i][k]);
          if (d) { if (o.indexOf('TU NGAY') === 0) ra.tu = d; else ra.den = d; break; }
        }
      }
    }
  }
  return ra;
}

/* doc thang bang BC TONG NO 3 CHI NHANH - so da doi chieu & dieu chinh khop APP */
function docBangChuan_(ss, tu, den, cn, kemCT) {
  var sh = layTab_(ss, TAB_TH);
  if (!sh) return { ok: false, loi: 'Khong tim thay tab "' + TAB_TH + '".' };
  var bang = sh.getDataRange().getDisplayValues();
  var h = timHangTieuDe_(bang, ['KHACH HANG', 'MA KH', 'CHI NHANH', 'NO DAU KY']);
  var hang = bang[h];
  var cTen = timCot_(hang, ['KHÁCH HÀNG']);
  var cMa  = timCot_(hang, ['MÃ KH']);
  var cCN  = timCot_(hang, ['CHI NHÁNH']);
  var cDK  = timCot_(hang, ['NỢ ĐẦU KỲ']);
  var cPS  = timCot_(hang, ['PHÁT SINH NỢ', 'PHÁT SINH']);
  var cThu = timCot_(hang, ['ĐÃ THU']);
  var cCK  = timCot_(hang, ['NỢ CUỐI KỲ']);
  var cApp = timCot_(hang, ['NỢ HIỆN TẠI APP']);
  var cDC  = timCot_(hang, ['ĐIỀU CHỈNH KHỚP APP']);
  var cGC  = timCot_(hang, ['CẢNH BÁO / GHI CHÚ', 'CẢNH BÁO']);
  if (cTen < 0 || cDK < 0 || cPS < 0 || cThu < 0) {
    return { ok: false, loi: 'Thieu cot trong tab "' + sh.getName() + '". Tieu de: ' + hang.join(' | ') };
  }

  var ds = [];
  for (var i = h + 1; i < bang.length; i++) {
    var r = bang[i];
    var ten = String(r[cTen] || '').trim();
    var stt = String(r[0] || '').trim();
    if (!ten) continue;
    if (khongDau_(stt).indexOf('TONG CONG') === 0) continue;
    if (!/^\d+$/.test(stt)) continue;                 /* chi lay dong co so thu tu */
    var g = {
      ten: ten,
      ma:  cMa  >= 0 ? String(r[cMa] || '').trim() : '',
      cn:  cCN  >= 0 ? String(r[cCN] || '').trim() : '',
      dauKy: soVN_(r[cDK]), ps: soVN_(r[cPS]), thu: soVN_(r[cThu]),
      app: cApp >= 0 ? soVN_(r[cApp]) : null,
      dc:  cDC  >= 0 ? soVN_(r[cDC])  : 0,
      gc:  cGC  >= 0 ? String(r[cGC] || '').trim() : ''
    };
    g.cuoiKy = cCK >= 0 ? soVN_(r[cCK]) : (g.dauKy + g.ps - g.thu);
    if (cn && khongDau_(g.cn) !== cn) continue;
    ds.push(g);
  }

  var ct = [];
  if (kemCT) ct = docChiTiet_(ss, tu, den, cn, ds);

  return {
    ok: true,
    nguon: 'BÁO CÁO NGÀY — tab ' + sh.getName() + ' (số đã đối chiếu & điều chỉnh khớp APP)',
    cheDo: 'BC',
    ky: { tu: inNgay_(tu), den: inNgay_(den) },
    chinhanh: cn || 'TẤT CẢ',
    tinhLuc: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss'),
    boQuaKhongNgay: 0,
    kh: ds,
    ct: ct
  };
}

/* chi tiet chung tu trong ky, lay tu SO CONG NO KH */
function docChiTiet_(ss, tu, den, cn, dsKH) {
  var sh = layTab_(ss, TAB_SO);
  if (!sh) return [];
  var soDong = sh.getLastRow(), soCot = sh.getLastColumn();
  if (soDong < 2) return [];
  var hang = sh.getRange(1, 1, 1, soCot).getDisplayValues()[0];
  var cNgay = timCot_(hang, ['NGÀY']);
  var cCNct = timCot_(hang, ['CN CHỨNG TỪ']);
  var cKH   = timCot_(hang, ['KHÁCH HÀNG']);
  var cCT   = timCot_(hang, ['CHỨNG TỪ']);
  var cDG   = timCot_(hang, ['DIỄN GIẢI']);
  var cPS   = timCot_(hang, ['PS NỢ', 'PHÁT SINH']);
  var cThu  = timCot_(hang, ['ĐÃ THU']);
  if (cNgay < 0 || cKH < 0 || cPS < 0 || cThu < 0) return [];

  var choPhep = null;
  if (cn && dsKH && dsKH.length) {
    choPhep = {};
    dsKH.forEach(function (x) { choPhep[khongDau_(x.ten)] = 1; });
  }
  var ngayVals = sh.getRange(2, cNgay + 1, soDong - 1, 1).getValues();
  var bang = sh.getRange(2, 1, soDong - 1, soCot).getDisplayValues();
  var tuMs = tu.getTime(), denMs = den.getTime(), ra = [];
  for (var i = 0; i < bang.length; i++) {
    var r = bang[i];
    var ten = String(r[cKH] || '').trim();
    if (!ten) continue;
    if (choPhep && !choPhep[khongDau_(ten)]) continue;
    var ps = soVN_(r[cPS]), thu = soVN_(r[cThu]);
    if (!ps && !thu) continue;
    var d = docNgay_(ngayVals[i][0]) || docNgay_(r[cNgay]);
    if (!d) continue;
    var ms = d.getTime();
    if (ms < tuMs || ms > denMs) continue;
    ra.push({
      ngay: inNgay_(d), kh: ten,
      cn: cCNct >= 0 ? String(r[cCNct] || '').trim() : '',
      ct: cCT >= 0 ? String(r[cCT] || '').trim() : '',
      dg: cDG >= 0 ? String(r[cDG] || '').trim() : '',
      ps: ps, thu: thu
    });
  }
  ra.sort(function (a, b) {
    var da = docNgay_(a.ngay), db = docNgay_(b.ngay);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });
  return ra;
}

/* ───────────────────────── chan doan ───────────────────────── */

function chanDoan_() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ds = ss.getSheets().map(function (s) { return s.getName(); });
  var ra = { ok: true, cacTab: ds, mau: {} };
  [TAB_SO, TAB_TH].forEach(function (ten) {
    var sh = layTab_(ss, ten);
    if (!sh) { ra.mau[ten] = 'KHONG TIM THAY'; return; }
    var r = Math.min(sh.getLastRow(), 9), c = Math.min(sh.getLastColumn(), 20);
    ra.mau[sh.getName()] = {
      soDong: sh.getLastRow(), soCot: sh.getLastColumn(),
      dongDau: sh.getRange(1, 1, r, c).getDisplayValues()
    };
  });
  return ra;
}

/* ───────────────────────── tinh toan ───────────────────────── */

function docDanhMucKH_(ss) {
  var map = {}, sh = layTab_(ss, TAB_TH);
  if (!sh) return map;
  var bang = sh.getDataRange().getDisplayValues();
  if (!bang.length) return map;
  var h = timHangTieuDe_(bang, ['KHACH HANG', 'MA KH', 'CHI NHANH', 'NO DAU KY']);
  var hang = bang[h];
  var cTen = timCot_(hang, ['KHÁCH HÀNG', 'TEN KHACH HANG']);
  var cMa  = timCot_(hang, ['MÃ KH', 'MA KHACH HANG']);
  var cCN  = timCot_(hang, ['CHI NHÁNH']);
  var cApp = timCot_(hang, ['NỢ HIỆN TẠI APP', 'NO HIEN TAI APP']);
  var cGC  = timCot_(hang, ['CẢNH BÁO / GHI CHÚ', 'CANH BAO', 'GHI CHÚ']);
  if (cTen < 0) return map;
  for (var i = h + 1; i < bang.length; i++) {
    var ten = String(bang[i][cTen] || '').trim();
    if (!ten) continue;
    if (khongDau_(String(bang[i][0] || '')).indexOf('TONG CONG') === 0) continue;
    if (khongDau_(ten).indexOf('TONG CONG') === 0) continue;
    var k = khongDau_(ten);
    if (map[k]) continue;
    map[k] = {
      ten: ten,
      ma:  cMa  >= 0 ? String(bang[i][cMa] || '').trim() : '',
      cn:  cCN  >= 0 ? String(bang[i][cCN] || '').trim() : '',
      app: cApp >= 0 ? soVN_(bang[i][cApp]) : null,
      gc:  cGC  >= 0 ? String(bang[i][cGC] || '').trim() : ''
    };
  }
  return map;
}

function tinhCongNo_(tu, den, cn, kemCT) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var dm = docDanhMucKH_(ss);

  var sh = layTab_(ss, TAB_SO);
  if (!sh) return { ok: false, loi: 'Khong tim thay tab "' + TAB_SO + '" trong file BAO CAO NGAY.' };

  var soDong = sh.getLastRow(), soCot = sh.getLastColumn();
  if (soDong < 2) return { ok: false, loi: 'Tab "' + sh.getName() + '" khong co du lieu.' };

  var hang = sh.getRange(1, 1, 1, soCot).getDisplayValues()[0];
  var cNgay = timCot_(hang, ['NGÀY']);
  var cCNct = timCot_(hang, ['CN CHỨNG TỪ', 'CHI NHÁNH']);
  var cKH   = timCot_(hang, ['KHÁCH HÀNG']);
  var cCT   = timCot_(hang, ['CHỨNG TỪ']);
  var cDG   = timCot_(hang, ['DIỄN GIẢI']);
  var cPS   = timCot_(hang, ['PS', 'PHÁT SINH']);
  var cThu  = timCot_(hang, ['ĐÃ THU', 'DA THU']);
  if (cNgay < 0 || cKH < 0 || cPS < 0 || cThu < 0) {
    return { ok: false, loi: 'Thieu cot bat buoc trong tab "' + sh.getName() + '". Tieu de doc duoc: ' + hang.join(' | ') };
  }

  var ngayVals = sh.getRange(2, cNgay + 1, soDong - 1, 1).getValues();
  var bang     = sh.getRange(2, 1, soDong - 1, soCot).getDisplayValues();

  var tuMs = tu.getTime(), denMs = den.getTime();
  var map = {}, thuTu = [], ct = [];
  var boQuaKhongNgay = 0;

  for (var i = 0; i < bang.length; i++) {
    var r = bang[i];
    var ten = String(r[cKH] || '').trim();
    if (!ten) continue;
    var ps = soVN_(r[cPS]), thu = soVN_(r[cThu]);
    if (!ps && !thu) continue;

    var d = docNgay_(ngayVals[i][0]) || docNgay_(r[cNgay]);
    if (!d) { boQuaKhongNgay++; continue; }
    var ms = d.getTime();
    if (ms > denMs) continue;

    var k = khongDau_(ten);
    if (!map[k]) {
      var info = dm[k] || {};
      map[k] = {
        ten: info.ten || ten, ma: info.ma || '', cn: info.cn || '',
        app: (info.app === undefined ? null : info.app), gc: info.gc || '',
        dauKy: 0, ps: 0, thu: 0, soCT: 0
      };
      thuTu.push(k);
    }
    var g = map[k];
    if (ms < tuMs) {
      g.dauKy += ps - thu;
    } else {
      g.ps  += ps;
      g.thu += thu;
      g.soCT++;
      if (kemCT) ct.push({
        ngay: inNgay_(d),
        kh: g.ten,
        cn: cCNct >= 0 ? String(r[cCNct] || '').trim() : '',
        ct: cCT >= 0 ? String(r[cCT] || '').trim() : '',
        dg: cDG >= 0 ? String(r[cDG] || '').trim() : '',
        ps: ps, thu: thu
      });
    }
  }

  var ds = [];
  for (var j = 0; j < thuTu.length; j++) {
    var g2 = map[thuTu[j]];
    if (cn) {
      var gcn = khongDau_(g2.cn);
      if (gcn !== cn) continue;
    }
    g2.cuoiKy = g2.dauKy + g2.ps - g2.thu;
    if (!g2.dauKy && !g2.ps && !g2.thu && !g2.cuoiKy) continue;
    ds.push(g2);
  }

  /* KH co phat sinh trong ky len truoc, con lai xep theo ten */
  ds.sort(function (a, b) {
    var pa = (a.ps + a.thu) > 0 ? 0 : 1, pb = (b.ps + b.thu) > 0 ? 0 : 1;
    if (pa !== pb) return pa - pb;
    return a.ten.localeCompare(b.ten, 'vi');
  });

  if (kemCT && cn) {
    var choPhep = {};
    ds.forEach(function (x) { choPhep[khongDau_(x.ten)] = 1; });
    ct = ct.filter(function (x) { return choPhep[khongDau_(x.kh)]; });
  }
  ct.sort(function (a, b) {
    var da = docNgay_(a.ngay), db = docNgay_(b.ngay);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });

  return {
    ok: true,
    nguon: 'BÁO CÁO NGÀY — tab ' + sh.getName() + ' (tính lại từ sổ cho kỳ đã chọn)',
    cheDo: 'SO',
    ky: { tu: inNgay_(tu), den: inNgay_(den) },
    chinhanh: cn || 'TẤT CẢ',
    tinhLuc: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss'),
    boQuaKhongNgay: boQuaKhongNgay,
    kh: ds,
    ct: kemCT ? ct : []
  };
}
