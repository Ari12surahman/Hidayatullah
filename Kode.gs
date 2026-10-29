function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle("Mutaba'ah Digital").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// 1. LOGIN ADMIN/MUSYRIF
function loginUser(u, p) {
  const s = getSheet('Data_User');
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() == String(u).trim() && String(d[i][1]).trim() == String(p).trim()) {
      return "SUKSES|" + d[i][2] + "|" + d[i][3] + "|" + d[i][0];
    }
  }
  return "GAGAL|Cek Username/Password";
}

// Helper membersihkan gelar / honorifik musyrif untuk pencocokan nama
function cleanMusyrifName(str) {
  if (!str) return "";
  var s = String(str).toLowerCase().trim();
  s = s.replace(/\b(ustadz|ustad|ust|guru|bpk|bapak|kyai|teungku|tengku|kak)\b\.?/gi, " ").trim();
  s = s.replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
  return s;
}

// Helper Format Nomor WA ke +62
function formatNomorWA(no) {
  if (!no) return "";
  var n = String(no).trim().replace(/[^0-9+]/g, "");
  if (!n) return "";
  if (n.startsWith("+62")) return n;
  if (n.startsWith("62")) return "+" + n;
  if (n.startsWith("0")) return "+62" + n.substring(1);
  if (n.startsWith("8")) return "+62" + n;
  return n;
}

// Helper untuk menterjemahkan identifier musyrif (bisa berupa username atau nama lengkap) ke dalam array alias
function getMusyrifAliases(identifier) {
  if (!identifier) return [];
  var identTrim = String(identifier).trim().toLowerCase();
  var aliases = [identTrim];
  var normIdent = cleanMusyrifName(identTrim);
  if (normIdent && aliases.indexOf(normIdent) === -1) aliases.push(normIdent);

  try {
    var s = getSheet('Data_User');
    if (s && s.getLastRow() >= 2) {
      var d = s.getRange(2, 1, s.getLastRow()-1, 4).getValues(); // col A=username, col C=nama, col D=role
      for (var i = 0; i < d.length; i++) {
        var usr = String(d[i][0] || '').trim().toLowerCase();
        var nm = String(d[i][2] || '').trim().toLowerCase();
        var normUsr = cleanMusyrifName(usr);
        var normNm = cleanMusyrifName(nm);

        var isMatch = (usr === identTrim || nm === identTrim || 
                       (normIdent && (normUsr === normIdent || normNm === normIdent)));
        if (isMatch) {
          if (usr && aliases.indexOf(usr) === -1) aliases.push(usr);
          if (nm && aliases.indexOf(nm) === -1) aliases.push(nm);
          if (normUsr && aliases.indexOf(normUsr) === -1) aliases.push(normUsr);
          if (normNm && aliases.indexOf(normNm) === -1) aliases.push(normNm);
        }
      }
    }
  } catch (e) {}
  return aliases;
}

function matchUserOrMusyrif(targetStr, aliases) {
  if (!targetStr || !aliases || !aliases.length) return false;
  var cleanTarget = String(targetStr).trim().toLowerCase();
  var normTarget = cleanMusyrifName(cleanTarget);
  if (!normTarget && !cleanTarget) return false;

  for (var i = 0; i < aliases.length; i++) {
    var a = String(aliases[i]).trim().toLowerCase();
    if (!a) continue;
    if (cleanTarget === a) return true;
    var normA = cleanMusyrifName(a);
    if (normA && normTarget) {
      if (normTarget === normA) return true;
      if (normTarget.length >= 4 && normA.length >= 4) {
        if (normTarget.indexOf(normA) !== -1 || normA.indexOf(normTarget) !== -1) return true;
      }
    }
  }
  return false;
}

// Master 114 Surat Al-Qur'an lengkap dan berurutan
var DAFTAR_114_SURAT = [
  "Al-Fatihah", "Al-Baqarah", "Ali 'Imran", "An-Nisa'", "Al-Ma'idah",
  "Al-An'am", "Al-A'raf", "Al-Anfal", "At-Taubah", "Yunus",
  "Hud", "Yusuf", "Ar-Ra'd", "Ibrahim", "Al-Hijr",
  "An-Nahl", "Al-Isra'", "Al-Kahf", "Maryam", "Taha",
  "Al-Anbiya'", "Al-Hajj", "Al-Mu'minun", "An-Nur", "Al-Furqan",
  "Asy-Syu'ara'", "An-Naml", "Al-Qasas", "Al-'Ankabut", "Ar-Rum",
  "Luqman", "As-Sajdah", "Al-Ahzab", "Saba'", "Fatir",
  "Yasin", "As-Saffat", "Sad", "Az-Zumar", "Ghafir",
  "Fussilat", "Asy-Syura", "Az-Zukhruf", "Ad-Dukhan", "Al-Jatsiyah",
  "Al-Ahqaf", "Muhammad", "Al-Fath", "Al-Hujurat", "Qaf",
  "Az-Zariyat", "At-Tur", "An-Najm", "Al-Qamar", "Ar-Rahman",
  "Al-Waqi'ah", "Al-Hadid", "Al-Mujadilah", "Al-Hasyr", "Al-Mumtahanah",
  "As-Saff", "Al-Jumu'ah", "Al-Munafiqun", "At-Taghabun", "At-Talaq",
  "At-Tahrim", "Al-Mulk", "Al-Qalam", "Al-Haqqah", "Al-Ma'arij",
  "Nuh", "Al-Jinn", "Al-Muzzammil", "Al-Muddatsir", "Al-Qiyamah",
  "Al-Insan", "Al-Mursalat", "An-Naba'", "An-Nazi'at", "'Abasa",
  "At-Takwir", "Al-Infitar", "Al-Muthaffifin", "Al-Insyiqaq", "Al-Buruj",
  "At-Tariq", "Al-A'la", "Al-Ghasyiyah", "Al-Fajr", "Al-Balad",
  "Asy-Syams", "Al-Lail", "Ad-Duha", "Asy-Syarh", "At-Tin",
  "Al-'Alaq", "Al-Qadr", "Al-Bayyinah", "Az-Zalzalah", "Al-'Adiyat",
  "Al-Qari'ah", "At-Takasur", "Al-'Asr", "Al-Humazah", "Al-Fil",
  "Quraisy", "Al-Ma'un", "Al-Kautsar", "Al-Kafirun", "An-Nasr",
  "Al-Lahab", "Al-Ikhlas", "Al-Falaq", "An-Nas"
];

// Pemetaan daftar surat per Juz untuk validasi dan filtering yang akurat
var SURAT_DALAM_JUZ = {
  30: [
    "An-Naba'", "An-Nazi'at", "'Abasa", "At-Takwir", "Al-Infitar", 
    "Al-Muthaffifin", "Al-Insyiqaq", "Al-Buruj", "At-Tariq", "Al-A'la", 
    "Al-Ghasyiyah", "Al-Fajr", "Al-Balad", "Asy-Syams", "Al-Lail", 
    "Ad-Duha", "Asy-Syarh", "At-Tin", "Al-'Alaq", "Al-Qadr", 
    "Al-Bayyinah", "Az-Zalzalah", "Al-'Adiyat", "Al-Qari'ah", "At-Takasur", 
    "Al-'Asr", "Al-Humazah", "Al-Fil", "Quraisy", "Al-Ma'un", 
    "Al-Kautsar", "Al-Kafirun", "An-Nasr", "Al-Lahab", "Al-Ikhlas", 
    "Al-Falaq", "An-Nas"
  ],
  29: [
    "Al-Mulk", "Al-Qalam", "Al-Haqqah", "Al-Ma'arij", "Nuh", 
    "Al-Jinn", "Al-Muzzammil", "Al-Muddatsir", "Al-Muddassir", "Al-Qiyamah", "Al-Insan", 
    "Al-Mursalat"
  ],
  28: [
    "Al-Mujadilah", "Al-Hasyr", "Al-Mumtahanah", "As-Saff", "Al-Jumu'ah", 
    "Al-Munafiqun", "At-Taghabun", "At-Talaq", "At-Tahrim"
  ]
};

// 2. AMBIL DATA SANTRI (WITH SERVER-SIDE SEARCH)
function getListSantri(u, r, searchQuery) {
  const s = getSheet('Data_Santri');
  if(s.getLastRow() < 2) return [];
  if(s.getMaxColumns() < 7) s.insertColumnsAfter(s.getMaxColumns(), 7 - s.getMaxColumns());
  
  // Pastikan header kolom E, F, G rapi
  try {
    var hdr = s.getRange(1, 1, 1, 7).getValues()[0];
    if (!hdr[4]) s.getRange(1, 5).setValue("Target");
    if (!hdr[5]) s.getRange(1, 6).setValue("No WA Wali");
    if (!hdr[6]) s.getRange(1, 7).setValue("Username Wali");
  } catch(e) {}

  // Ambil 7 Kolom: Nama, Kelas, Musyrif, Password, Target, No WA Wali, Username Wali
  const numCols = Math.max(7, s.getLastColumn());
  const d = s.getRange(2, 1, s.getLastRow()-1, numCols).getValues();
  
  let allData = d.map((row, i) => ({
    nama: String(row[0] || '').trim(), 
    kelas: String(row[1] || '').trim(), 
    musyrif: String(row[2] || '').trim(), 
    password: String(row[3] || '').trim(),
    target: String(row[4] || '').trim(), 
    noWaWali: String(row[5] || '').trim(),
    usernameWali: String(row[6] || '').trim(),
    row: i + 2
  })).filter(x => x.nama);

  var roleSafety = String(r || '').toLowerCase();
  // KEAMANAN FILTER: Hanya Admin yang melihat seluruh santri. Musyrif hanya melihat santri binaannya sendiri.
  if(roleSafety !== 'admin') {
    var uAliases = getMusyrifAliases(u);
    allData = allData.filter(x => matchUserOrMusyrif(x.musyrif, uAliases));
  }

  // Server-side search filter
  if (searchQuery && searchQuery.trim() !== "") {
    let q = searchQuery.toLowerCase();
    allData = allData.filter(x => 
      String(x.nama).toLowerCase().includes(q) || 
      String(x.kelas).toLowerCase().includes(q) ||
      String(x.musyrif).toLowerCase().includes(q) ||
      String(x.noWaWali).toLowerCase().includes(q) ||
      String(x.usernameWali).toLowerCase().includes(q)
    );
  }

  // Kalkulasi Progress Hafalan untuk Setiap Santri
  var masterKelasDict = {};
  try {
    var shKelas = getSheet('Master_Kelas');
    if(shKelas && shKelas.getLastRow() >= 2) {
      var kData = shKelas.getRange(2, 1, shKelas.getLastRow()-1, 2).getValues();
      kData.forEach(k => { if(k[0]) masterKelasDict[String(k[0]).trim().toLowerCase()] = String(k[1] || '').trim(); });
    }
  } catch(e) {}

  var mqList = getMasterQuran();
  var targetList = getTargetHafalan();
  var santriTargets = {};
  var kelasTargets = {};
  targetList.forEach(t => {
    if (String(t.tipe).toLowerCase() === 'santri' && t.targetUntuk) {
      santriTargets[String(t.targetUntuk).trim().toLowerCase()] = String(t.detail || '').trim();
    } else if (String(t.tipe).toLowerCase() === 'kelas' && t.targetUntuk) {
      kelasTargets[String(t.targetUntuk).trim().toLowerCase()] = String(t.detail || '').trim();
    }
  });

  var santriStats = {}; 
  try {
    var shLapor = getSheet('Laporan');
    if (shLapor && shLapor.getLastRow() >= 2) {
      var lapData = shLapor.getRange(2, 4, shLapor.getLastRow()-1, 7).getValues();
      lapData.forEach(r => {
        var nm = String(r[0]).trim();
        if (!nm) return;
        if (!santriStats[nm]) santriStats[nm] = { surahs: [], hal: 0, ayat: 0, count: 0, rangesPerSurat: {}, uniquePages: {}, setoranList: [] };
        var stats = santriStats[nm];
        stats.count++;
        var srt = String(r[2]).trim();
        var rowJenis = String(r[1]).trim().toUpperCase();
        if (srt && srt !== '-' && srt !== '') {
          if (stats.surahs.indexOf(srt) === -1) stats.surahs.push(srt);
        }
        var hAwal = parseInt(r[3]) || 0;
        var hAkhir = parseInt(r[4]) || hAwal;
        if (hAwal === 0 && hAkhir === 0 && srt.toLowerCase().indexOf('juz') !== -1) {
          var mJ = srt.match(/juz\s*(\d+)/i);
          if (mJ && mJ[1]) {
            var jR = getJuzRange(parseInt(mJ[1]));
            if (jR) { hAwal = jR.start; hAkhir = jR.end; }
          }
        }
        if (hAwal > 0 || hAkhir > 0) {
          for (var pg = Math.min(hAwal, hAkhir); pg <= Math.max(hAwal, hAkhir); pg++) {
            stats.uniquePages[pg] = true;
          }
        }
        
        var aAwal = parseInt(r[5]) || 0;
        var aAkhir = parseInt(r[6]) || aAwal;
        stats.setoranList.push({
          jenis: rowJenis,
          surat: srt,
          halAwal: hAwal,
          halAkhir: hAkhir,
          ayatAwal: aAwal,
          ayatAkhir: aAkhir
        });

        // Track ayat ranges per surat untuk Ziyadah (hitung unik)
        if (rowJenis === 'Z' && srt && aAwal > 0 && aAkhir > 0) {
          if (!stats.rangesPerSurat[srt]) stats.rangesPerSurat[srt] = [];
          stats.rangesPerSurat[srt].push([aAwal, aAkhir]);
        } else {
          // Non-Ziyadah: hitung biasa
          if (aAwal || aAkhir) stats.ayat += Math.abs(aAkhir - aAwal) + 1;
          else if (stats.count) stats.ayat += 10;
        }
      });
      // Hitung ayat unik dari Ziyadah ranges & halaman unik
      for (var nmKey in santriStats) {
        var st = santriStats[nmKey];
        st.hal = Object.keys(st.uniquePages || {}).length;
        for (var surat in st.rangesPerSurat) {
          var merged = mergeRanges(st.rangesPerSurat[surat]);
          st.ayat += hitungTotalAyatDariRanges(merged);
        }
      }
    }
  } catch(e) {}

  allData.forEach(x => {
    var nmLower = String(x.nama || '').trim().toLowerCase();
    var klsLower = String(x.kelas || '').trim().toLowerCase();
    
    var targetStr = santriTargets[nmLower] || String(x.target || '').trim();
    if (!targetStr && klsLower) {
      targetStr = kelasTargets[klsLower] || masterKelasDict[klsLower] || '';
    }
    
    var stats = santriStats[String(x.nama).trim()] || { surahs: [], hal: 0, ayat: 0, count: 0, setoranList: [] };
    x.progress = kalkulasiProgressTarget(targetStr, stats, mqList);
  });

  return allData;
}

// --- HELPER KALKULASI PROGRES TERSTRUKTUR ---
function getJuzRange(juzNum) {
  var j = parseInt(juzNum);
  if (j === 1) return { start: 1, end: 21, totalHal: 21, nama: "Juz 1" };
  if (j === 30) return { start: 582, end: 604, totalHal: 23, nama: "Juz 30" };
  if (j >= 2 && j <= 29) {
    var s = ((j - 1) * 20) + 2;
    var e = (j * 20) + 1;
    return { start: s, end: e, totalHal: 20, nama: "Juz " + j };
  }
  return null;
}

// Helper untuk menganalisis target Juz: mendukung single juz (Juz 29), rentang juz (Juz 29 s/d 30, Juz 28-30), maupun jumlah juz (2 Juz)
function getMultiJuzInfo(str) {
  if (!str) return null;
  var s = String(str).trim();
  var sLower = s.toLowerCase();

  // 1. Cek rentang juz: "Juz 29 s/d 30", "Juz 29 - 30", "Juz 28 sampai 30", "Juz 1 s/d 3", dll
  var mRange = sLower.match(/juz\s*(\d+)\s*(?:s\/d|-|sampai|hingga|&|,|\+)\s*(?:juz\s*)?(\d+)/i);
  if (mRange && mRange[1] && mRange[2]) {
    var j1 = parseInt(mRange[1]);
    var j2 = parseInt(mRange[2]);
    var startJ = Math.min(j1, j2);
    var endJ = Math.max(j1, j2);
    var allPages = {};
    var totalHal = 0;
    var surahsInMulti = {};

    for (var j = startJ; j <= endJ; j++) {
      var r = getJuzRange(j);
      if (r) {
        totalHal += r.totalHal;
        for (var p = r.start; p <= r.end; p++) allPages[p] = true;
      }
      if (typeof SURAT_DALAM_JUZ !== 'undefined' && SURAT_DALAM_JUZ[j]) {
        SURAT_DALAM_JUZ[j].forEach(function(srt) { surahsInMulti[normalizeSurat(srt)] = true; });
      }
    }
    return {
      type: 'range',
      startJuz: startJ,
      endJuz: endJ,
      totalHal: totalHal,
      pages: allPages,
      suratMap: surahsInMulti,
      nama: "Juz " + startJ + " s/d " + endJ
    };
  }

  // 2. Cek jumlah Juz umum: "2 Juz", "3 Juz", "5 Juz"
  var mCount = sLower.match(/(\d+)\s*juz/i);
  if (mCount && mCount[1] && sLower.indexOf("juz 2") !== 0 && sLower.indexOf("juz 1") !== 0 && sLower.indexOf("juz 3") !== 0) {
    var numCount = parseInt(mCount[1]);
    if (numCount > 0 && numCount <= 30) {
      return {
        type: 'count',
        totalHal: numCount * 20,
        nama: numCount + " Juz",
        pages: null,
        suratMap: {}
      };
    }
  }

  // 3. Single Juz: "Juz 29", "JUZ:29", "Juz 30", dll
  var mSingle = s.match(/JUZ:(\d+)/i) || sLower.match(/juz\s*(\d+)/i);
  if (mSingle && mSingle[1]) {
    var numJuz = parseInt(mSingle[1]);
    var r = getJuzRange(numJuz);
    if (r) {
      var allPages = {};
      for (var p = r.start; p <= r.end; p++) allPages[p] = true;
      var surahsSingle = {};
      if (typeof SURAT_DALAM_JUZ !== 'undefined' && SURAT_DALAM_JUZ[numJuz]) {
        SURAT_DALAM_JUZ[numJuz].forEach(function(srt) { surahsSingle[normalizeSurat(srt)] = true; });
      }
      return {
        type: 'single',
        startJuz: numJuz,
        endJuz: numJuz,
        totalHal: r.totalHal,
        pages: allPages,
        suratMap: surahsSingle,
        nama: r.nama
      };
    }
  }

  return null;
}

function kalkulasiProgressTarget(targetStr, stats, mqList) {
  var srtCount = stats.surahs ? stats.surahs.length : 0;
  var halCount = stats.hal || 0;
  var ayatCount = stats.ayat || 0;
  var setoranList = stats.setoranList || [];
  
  var persen = 0;
  var tercapai = "";
  var sisa = "";
  var tTrim = String(targetStr || '').trim();
  var tLower = tTrim.toLowerCase();
  
  if (!tTrim) {
    persen = Math.min(100, Math.round(((stats.count || 0) / 20) * 100));
    tercapai = srtCount + " Surat (" + halCount + " Hal)";
    sisa = "Belum ada target khusus dirinci";
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 1. TERSTRUKTUR / HEURISTIK JUZ (MENDUKUNG SINGLE JUZ MAUPUN MULTI-JUZ SEPERTI JUZ 29 S/D 30)
  if (tTrim.indexOf("JUZ:") === 0 || tLower.indexOf("juz") !== -1) {
    var juzInfo = getMultiJuzInfo(tTrim);
    if (juzInfo) {
      // Gabungkan daftar surat dari Master_Quran sheet + pemetaan bawaan SURAT_DALAM_JUZ
      var suratDalamJuz = Object.assign({}, juzInfo.suratMap || {});
      if (mqList && mqList.length > 0 && juzInfo.pages) {
        mqList.forEach(function(m) {
          if (m.halaman && juzInfo.pages[m.halaman] && m.surat) {
            suratDalamJuz[normalizeSurat(m.surat)] = true;
          }
        });
      }
      
      var pagesDoneInTarget = {};
      var filteredSurahs = [];

      setoranList.forEach(function(item) {
        var itemSuratNorm = item.surat ? normalizeSurat(item.surat) : '';
        var isInTargetJuz = false;
        
        // Cek berdasarkan rentang halaman setoran
        if (juzInfo.pages && (item.halAwal > 0 || item.halAkhir > 0)) {
          var ha = item.halAwal || item.halAkhir;
          var hk = item.halAkhir || item.halAwal;
          for (var p = Math.min(ha, hk); p <= Math.max(ha, hk); p++) {
            if (juzInfo.pages[p]) {
              pagesDoneInTarget[p] = true;
              isInTargetJuz = true;
            }
          }
        }
        
        // Cek berdasarkan nama surat jika cocok dengan surat di juz target
        if (!isInTargetJuz && itemSuratNorm && suratDalamJuz[itemSuratNorm]) {
          isInTargetJuz = true;
        }
        
        if (isInTargetJuz && item.surat && item.surat !== '-' && item.surat !== '') {
          if (filteredSurahs.indexOf(item.surat) === -1) filteredSurahs.push(item.surat);
        }
      });
      
      var completedHalInTarget = Object.keys(pagesDoneInTarget).length;
      var filteredSrtCount = filteredSurahs.length;
      var targetTotalHal = juzInfo.totalHal || 20;

      // Hitung halaman ekstra santri di luar target ini (misal santri sudah tuntas Juz 29 dan juga punya hafalan Juz 30 / Juz 28)
      var totalSantriHal = halCount;
      var extraHalDiLuarTarget = Math.max(0, totalSantriHal - completedHalInTarget);
      var totalJuzEst = (totalSantriHal / 20).toFixed(1);

      // Jika setoran tanpa nomor halaman (hanya nama surat)
      if (completedHalInTarget === 0 && filteredSrtCount > 0 && juzInfo.type === 'single') {
        var maxS = (juzInfo.startJuz === 30) ? 37 : ((juzInfo.startJuz === 29) ? 11 : 10);
        persen = Math.min(100, Math.round((filteredSrtCount / maxS) * 100));
        tercapai = filteredSrtCount + " dari " + maxS + " Surat (" + juzInfo.nama + ")";
        if (filteredSrtCount >= maxS) {
          sisa = "Target " + juzInfo.nama + " Tuntas 100%! 🎉" + (extraHalDiLuarTarget > 0 ? " (+ hafalan lain " + extraHalDiLuarTarget + " hal)" : " Siap lanjut target grade berikutnya 🚀");
        } else {
          sisa = "Sisa " + (maxS - filteredSrtCount) + " surat lagi di " + juzInfo.nama;
        }
      } else {
        // Berdasarkan halaman mushaf yang tuntas
        persen = Math.min(100, Math.round((completedHalInTarget / targetTotalHal) * 100));
        
        if (completedHalInTarget >= targetTotalHal) {
          persen = 100;
          tercapai = juzInfo.nama + " Tuntas (100%)" + (extraHalDiLuarTarget > 0 ? " • Total: " + totalSantriHal + " Hal (~" + totalJuzEst + " Juz)" : " (" + completedHalInTarget + " Hal)");
          sisa = "Target " + juzInfo.nama + " Selesai! 🎉" + (extraHalDiLuarTarget > 0 ? " (+ ada " + extraHalDiLuarTarget + " hal di luar target). Siap perbarui target grade! 🚀" : " Siap lanjut target grade berikutnya 🚀");
        } else {
          var sisaH = Math.max(0, targetTotalHal - completedHalInTarget);
          tercapai = completedHalInTarget + " dari " + targetTotalHal + " Hal (" + persen + "%)" + (extraHalDiLuarTarget > 0 ? " • Total: " + totalSantriHal + " Hal" : "");
          sisa = "Sisa " + sisaH + " halaman lagi di " + juzInfo.nama + (extraHalDiLuarTarget > 0 ? " (+ sudah ada " + extraHalDiLuarTarget + " hal di juz lain)" : "");
        }
      }

      return { 
        persen: persen, 
        tercapai: tercapai, 
        sisa: sisa, 
        surat: filteredSrtCount || srtCount, 
        hal: completedHalInTarget || totalSantriHal, 
        totalSantriHal: totalSantriHal,
        ayat: ayatCount 
      };
    }

    var maxHal = 600;
    persen = Math.min(100, Math.round((halCount / maxHal) * 100));
    tercapai = halCount + " dari " + maxHal + " Halaman (" + srtCount + " Surat)";
    sisa = halCount >= maxHal ? "Target 30 Juz Selesai! 🎉" : ("Sisa " + Math.max(0, maxHal - halCount) + " halaman lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 2. TERSTRUKTUR SURAT / RENTANG SURAT
  if (tTrim.indexOf("SURAT:") === 0 || tLower.indexOf("surat") !== -1 || tLower.indexOf("surah") !== -1) {
    var parts = tTrim.split("|");
    if (tTrim.indexOf("SURAT:") === 0 && parts.length >= 3 && mqList && mqList.length > 0) {
      var sStart = parts[1].trim();
      var sEnd = parts[2].split(" ")[0].trim();
      var pStart = 604, pEnd = 1;
      mqList.forEach(function(m) {
        if (isSuratMatch(m.surat, sStart)) { if (m.halaman < pStart) pStart = m.halaman; }
        if (isSuratMatch(m.surat, sEnd)) { if (m.halaman > pEnd) pEnd = m.halaman; }
      });
      if (pStart <= pEnd) {
        var totalH = (pEnd - pStart) + 1;
        var pagesDone = {};
        setoranList.forEach(function(item) {
          if (item.halAwal > 0 || item.halAkhir > 0) {
            for (var p = Math.min(item.halAwal, item.halAkhir); p <= Math.max(item.halAwal, item.halAkhir); p++) {
              if (p >= pStart && p <= pEnd) pagesDone[p] = true;
            }
          }
        });
        var completed = Object.keys(pagesDone).length;
        persen = Math.min(100, Math.round((completed / totalH) * 100));
        tercapai = completed + " dari " + totalH + " Halaman (" + sStart + " s/d " + sEnd + ")";
        var sisaH = Math.max(0, totalH - completed);
        sisa = sisaH > 0 ? ("Sisa " + sisaH + " halaman lagi untuk rentang surah ini") : "Target Rentang Surah Selesai! 🎉";
        return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: completed, ayat: ayatCount };
      }
    }
    var matchS = tLower.match(/(\d+)/);
    var maxS = matchS ? parseInt(matchS[1]) : 10;
    persen = Math.min(100, Math.round((srtCount / maxS) * 100));
    tercapai = srtCount + " dari " + maxS + " Surat";
    sisa = srtCount >= maxS ? "Target Surat Selesai! 🎉" : ("Sisa " + (maxS - srtCount) + " surat lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 3. TERSTRUKTUR HALAMAN / TAHAP
  if (tTrim.indexOf("HAL:") === 0 || tLower.indexOf("hal") !== -1 || tLower.indexOf("lembar") !== -1 || tLower.indexOf("hlm") !== -1) {
    var matchHal = tTrim.match(/HAL:(\d+)/i) || tLower.match(/(\d+)/);
    var maxH = matchHal ? parseInt(matchHal[1]) : 10;
    if (tLower.indexOf("lembar") !== -1) maxH = maxH * 2;
    persen = Math.min(100, Math.round((halCount / maxH) * 100));
    tercapai = halCount + " dari " + maxH + " Halaman";
    sisa = halCount >= maxH ? "Target Halaman Selesai! 🎉" : ("Sisa " + (maxH - halCount) + " halaman lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  if (tTrim.indexOf("TAHAP:") === 0 || tLower.indexOf("tahap") !== -1 || tLower.indexOf("jilid") !== -1 || tLower.indexOf("iqro") !== -1) {
    var maxJ = 30;
    persen = Math.min(100, Math.round((halCount / maxJ) * 100));
    tercapai = halCount + " Halaman / Bab tuntas";
    sisa = halCount >= maxJ ? "Target Jilid/Tahap Selesai! 🎉" : ("Sisa " + (maxJ - halCount) + " halaman/latihan lagi");
    return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
  }

  // 4. GENERAL / LAINNYA (Standar estimasi 1 Juz = 20 Hal / 300 Ayat)
  var maxH = 20;
  var maxA = 300;
  persen = halCount > 0 ? Math.min(100, Math.round((halCount / maxH) * 100)) : Math.min(100, Math.round((ayatCount / maxA) * 100));
  tercapai = srtCount + " Surat (" + halCount + " Hal / " + ayatCount + " Ayat)";
  var sisaH = Math.max(0, maxH - halCount);
  var sisaA = Math.max(0, maxA - ayatCount);
  sisa = (halCount >= maxH || ayatCount >= maxA) ? "Target Selesai / Tercapai! 🎉" : ("Sisa " + sisaH + " Halaman (" + sisaA + " Ayat) lagi");
  return { persen: persen, tercapai: tercapai, sisa: sisa, surat: srtCount, hal: halCount, ayat: ayatCount };
}

function getProgressSantriSingle(nama, targetStr) {
  try {
    var all = getListSantri('', 'admin');
    var found = all.find(s => String(s.nama).trim() === String(nama).trim());
    if (found && found.progress) return found.progress;
  } catch(e) {}
  return { persen: 0, tercapai: "0 Surat (0 Hal)", sisa: "Belum ada data setoran", surat: 0, hal: 0, ayat: 0 };
}

function getListSurat() {
  try {
    const s = getSheet('Data_Surat');
    // Jika sheet Data_Surat kosong atau belum memuat lengkap 114 surat, sinkronkan otomatis
    if (s.getLastRow() < 115) {
      s.clearContents();
      s.getRange(1, 1).setValue("Nama Surat");
      var rows = DAFTAR_114_SURAT.map(function(srt, idx) {
        return [(idx + 1) + ". " + srt];
      });
      s.getRange(2, 1, rows.length, 1).setValues(rows);
      return rows.map(function(r) { return r[0]; });
    }
    var vals = s.getRange(2, 1, s.getLastRow()-1, 1).getValues().map(x => String(x[0]).trim()).filter(n => n);
    if (vals.length < 114) {
      return DAFTAR_114_SURAT.map(function(srt, idx) { return (idx + 1) + ". " + srt; });
    }
    return vals;
  } catch(e) {
    return DAFTAR_114_SURAT.map(function(srt, idx) { return (idx + 1) + ". " + srt; });
  }
}

function getListMusyrif() {
  const s = getSheet('Data_User');
  if(s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow()-1, 4).getValues()
    .map((r,i)=>({username:r[0], nama:r[2], role:r[3]}))
    .filter(u => String(u.role).toLowerCase() === 'musyrif');
}

// BUNDLED INITIAL DATA FOR PERFORMANCE
function getInitialData(u, r) {
  try {
    return {
      settings: getAppSettings(),
      santri: getListSantri(u, r),
      surahs: getListSurat(),
      targets: getTargetHafalan(),
      masterKelas: getMasterKelas(),
      masterQuran: getMasterQuran()
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// 3. SIMPAN LAPORAN (dengan deteksi overlap untuk Ziyadah)
function prosesSimpan(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setSpreadsheetTimeZone("Asia/Jakarta"); 
    
    var formattedDate = "";
    if(data.tglManual) {
        var parts = data.tglManual.split('-'); 
        formattedDate = parts[2] + "/" + parts[1] + "/" + parts[0]; 
    } else {
        formattedDate = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
    }

    // Auto-fill halaman dari Master_Quran jika belum diisi
    var halAwal = data.halAwal || '-';
    var halAkhir = data.halAkhir || '-';
    if ((halAwal === '-' || halAwal === '') && data.surat && data.ayatAwal && data.ayatAwal !== '-') {
      var hasilHal = cariHalamanByAyat(data.surat, data.ayatAwal, data.ayatAkhir);
      if (hasilHal) {
        halAwal = String(hasilHal.halAwal);
        halAkhir = String(hasilHal.halAkhir);
      }
    }

    // Cek overlap untuk Ziyadah SEBELUM simpan
    var overlapInfo = null;
    if (String(data.jenis).toUpperCase() === 'Z') {
      overlapInfo = cekOverlapZiyadah(data.nama, data.surat, data.ayatAwal, data.ayatAkhir, halAwal, halAkhir);
    }

    // Selalu simpan ke Laporan (riwayat tetap tercatat)
    var catatan = data.catatan || '';
    if (overlapInfo && overlapInfo.status === 'FULL_OVERLAP') {
      catatan = (catatan ? catatan + ' | ' : '') + '[Pengulangan - tidak menambah progres]';
    } else if (overlapInfo && overlapInfo.status === 'PARTIAL') {
      catatan = (catatan ? catatan + ' | ' : '') + '[Sebagian baru: +' + overlapInfo.ayatBaru + ' ayat]';
    }

    getSheet('Laporan').appendRow([
      formattedDate,    
      data.jamMulai,    
      data.jamSelesai,  
      data.nama,        
      data.jenis,       
      data.surat,       
      halAwal,
      halAkhir,
      data.ayatAwal || '-',
      data.ayatAkhir || '-',
      data.nilai,       
      catatan,     
      data.pelapor      
    ]);
    
    // Return object dengan info overlap
    return {
      status: "SUKSES",
      overlapInfo: overlapInfo,
      halAwal: halAwal,
      halAkhir: halAkhir
    };
  } catch(e) { return { status: "ERROR", message: e.toString() }; }
}

function prosesSimpanBatch(dataList) {
  try {
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     ss.setSpreadsheetTimeZone("Asia/Jakarta");
     var sheet = getSheet('Laporan');
     var rows = [];
     var batchOverlapResults = [];

     // dataList is now an array of full data objects
     dataList.forEach(function(data) {
        var formattedDate = "";
        if (data.tglManual) {
            var parts = data.tglManual.split('-'); 
            formattedDate = parts[2] + "/" + parts[1] + "/" + parts[0]; 
        } else {
            formattedDate = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
        }

        // Auto-fill halaman dari Master_Quran jika belum diisi
        var halAwal = data.halAwal || '-';
        var halAkhir = data.halAkhir || '-';
        if ((halAwal === '-' || halAwal === '') && data.surat && data.ayatAwal && data.ayatAwal !== '-') {
          var hasilHal = cariHalamanByAyat(data.surat, data.ayatAwal, data.ayatAkhir);
          if (hasilHal) {
            halAwal = String(hasilHal.halAwal);
            halAkhir = String(hasilHal.halAkhir);
          }
        }

        // Cek overlap untuk Ziyadah
        var catatan = data.catatan || '';
        if (String(data.jenis).toUpperCase() === 'Z') {
          var overlapInfo = cekOverlapZiyadah(data.nama, data.surat, data.ayatAwal, data.ayatAkhir, halAwal, halAkhir);
          if (overlapInfo) {
            batchOverlapResults.push({ nama: data.nama, surat: data.surat, info: overlapInfo });
            if (overlapInfo.status === 'FULL_OVERLAP') {
              catatan = (catatan ? catatan + ' | ' : '') + '[Pengulangan - tidak menambah progres]';
            } else if (overlapInfo.status === 'PARTIAL') {
              catatan = (catatan ? catatan + ' | ' : '') + '[Sebagian baru: +' + overlapInfo.ayatBaru + ' ayat]';
            }
          }
        }

        rows.push([
          formattedDate,
          data.jamMulai,
          data.jamSelesai,
          data.nama,
          data.jenis,
          data.surat,
          halAwal,
          halAkhir,
          data.ayatAwal || '-',
          data.ayatAkhir || '-',
          data.nilai,
          catatan,
          data.pelapor
        ]);
     });
     
     if(rows.length > 0) {
       // Batch insert
       sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
     }
     
     return {
       status: "SUKSES",
       count: rows.length,
       overlapResults: batchOverlapResults
     };
  } catch(e) { return { status: "ERROR", message: e.toString() }; }
}

// 4. RIWAYAT DENGAN PAGINATION & SEARCH
function getRiwayat(role, user, tglFilter, page, limit, searchQuery) {
  const s = getSheet('Laporan');
  if(s.getLastRow() < 2) return { data: [], total: 0, page: 1, totalPages: 0 };
  
  // Set default page dan limit
  page = page || 1;
  limit = limit || 10;
  
  const rawData = s.getDataRange().getDisplayValues();
  rawData.shift(); 

  const sheetSantri = getSheet('Data_Santri');
  let mapKelas = {}; 
  let mapSantriMusyrif = {};
  if (sheetSantri.getLastRow() >= 2) {
    let dataSantri = sheetSantri.getRange(2, 1, sheetSantri.getLastRow()-1, 3).getValues();
    dataSantri.forEach(r => { 
      var sName = String(r[0] || '').trim();
      if (sName) {
        mapKelas[sName] = String(r[1] || '').trim();
        mapSantriMusyrif[sName] = String(r[2] || '').trim();
      }
    });
  }
  
  // Map dengan index untuk mendapatkan row number
  let indexedData = rawData.map((r, idx) => ({ row: idx + 2, data: r }));
  
  var roleSafety = String(role || '').toLowerCase();
  var isAdmin = (roleSafety === 'admin');
  var userAliases = !isAdmin ? getMusyrifAliases(user) : [];
  let filtered = indexedData.filter(item => {
    var r = item.data;
    var tglSheet = r[0]; 
    var pecah = tglSheet.split('/');   
    var tglISO = "";
    if(pecah.length === 3) {
       tglISO = pecah[2] + "-" + pecah[1] + "-" + pecah[0];
    } else {
       tglISO = tglSheet;
    }
    
    let matchTgl = true;
    if (tglFilter) matchTgl = (tglISO === tglFilter);
    
    let matchUser = true;
    if(!isAdmin) {
      var sName = String(r[3] || '').trim();
      var sMusyrif = mapSantriMusyrif[sName] || '';
      var sPelapor = String(r[12] || '').trim();
      matchUser = matchUserOrMusyrif(sMusyrif, userAliases) || matchUserOrMusyrif(sPelapor, userAliases);
    }

    let matchSearch = true;
    if (searchQuery && searchQuery.trim() !== "") {
        let q = searchQuery.toLowerCase();
        let nama = String(r[3]).toLowerCase();
        let surat = String(r[5]).toLowerCase();
        let jenis = String(r[4]).toLowerCase();
        let nilai = String(r[10]).toLowerCase();
        let catatan = String(r[11]).toLowerCase();
        
        matchSearch = nama.includes(q) || surat.includes(q) || jenis.includes(q) || nilai.includes(q) || catatan.includes(q);
    }

    return matchTgl && matchUser && matchSearch;
  });

  let allData = filtered.map(item => {
    var r = item.data;
    // Convert tanggal ke format ISO untuk input date
    var tglSheet = r[0];
    var tglISO = "";
    var pecah = tglSheet.split('/');
    if(pecah.length === 3) {
      tglISO = pecah[2] + "-" + pecah[1].padStart(2,'0') + "-" + pecah[0].padStart(2,'0');
    } else {
      tglISO = tglSheet;
    }
    
    return {
      row: item.row,
      tanggal: r[0],
      tanggalISO: tglISO,
      mulai: r[1],
      selesai: r[2],
      waktu: r[1] + " - " + r[2], 
      nama: r[3], 
      kelas: mapKelas[r[3]] || '-', 
      jenis: r[4], 
      surat: r[5], 
      halAwal: r[6],
      halAkhir: r[7],
      ayatAwal: r[8],
      ayatAkhir: r[9],
      nilai: r[10], 
      catatan: r[11], 
      pelapor: r[12]
    };
  });
  
  // Reverse untuk data terbaru di atas
  allData = allData.reverse();
  
  // Hitung pagination
  var total = allData.length;
  var totalPages = Math.ceil(total / limit);
  var startIndex = (page - 1) * limit;
  var endIndex = startIndex + limit;
  var pagedData = allData.slice(startIndex, endIndex);
  
  return {
    data: pagedData,
    total: total,
    page: page,
    totalPages: totalPages
  };
}

// 4b. UPDATE LAPORAN
function updateLaporan(rowNum, data) {
  try {
    var sheet = getSheet('Laporan');
    
    // Format tanggal dari ISO ke dd/MM/yyyy
    var formattedDate = "";
    if(data.tglManual) {
      var parts = data.tglManual.split('-'); 
      formattedDate = parts[2] + "/" + parts[1] + "/" + parts[0]; 
    }
    
    // Auto-fill halaman dari Master_Quran jika belum diisi atau kosong
    var halAwal = data.halAwal || '-';
    var halAkhir = data.halAkhir || '-';
    if ((halAwal === '-' || halAwal === '') && data.surat && data.ayatAwal && data.ayatAwal !== '-') {
      var hasilHal = cariHalamanByAyat(data.surat, data.ayatAwal, data.ayatAkhir);
      if (hasilHal) {
        halAwal = String(hasilHal.halAwal);
        halAkhir = String(hasilHal.halAkhir);
      }
    }

    // Update row: Tanggal, Mulai, Selesai, Nama, Jenis, Surat, HalAwal, HalAkhir, AyatAwal, AyatAkhir, Nilai, Catatan, Musyrif
    sheet.getRange(rowNum, 1, 1, 13).setValues([[
      formattedDate,
      data.jamMulai,
      data.jamSelesai,
      data.nama,
      data.jenis,
      data.surat,
      halAwal,
      halAkhir,
      data.ayatAwal || '-',
      data.ayatAkhir || '-',
      data.nilai,
      data.catatan,
      data.pelapor
    ]]);
    
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// 4c. HAPUS LAPORAN
function hapusLaporan(rowNum) {
  try {
    var sheet = getSheet('Laporan');
    sheet.deleteRow(rowNum);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// 4d. HAPUS BULK LAPORAN
function deleteBulkSetoran(rowIds) {
  try {
    var sheet = getSheet('Laporan');
    // Sort descending to avoid index shifting
    rowIds.sort(function(a, b){return b-a});
    
    for(var i=0; i<rowIds.length; i++){
       sheet.deleteRow(rowIds[i]);
    }
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// Helper konversi tanggal ke ISO YYYY-MM-DD berpad 0 agar perbandingan akurat
function convertToISODate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1).padStart(2, '0');
    var d = String(val.getDate()).padStart(2, '0');
    return y + "-" + m + "-" + d;
  }
  var str = String(val).trim();
  if (str.includes('/')) {
    var p = str.split('/');
    if (p.length === 3) {
      if (p[0].length === 4) {
        return p[0] + "-" + String(p[1]).padStart(2, '0') + "-" + String(p[2]).padStart(2, '0');
      } else {
        return p[2] + "-" + String(p[1]).padStart(2, '0') + "-" + String(p[0]).padStart(2, '0');
      }
    }
  } else if (str.includes('-')) {
    var p = str.split('-');
    if (p.length === 3) {
      if (p[0].length === 4) {
        return p[0] + "-" + String(p[1]).padStart(2, '0') + "-" + String(p[2]).padStart(2, '0');
      } else {
        return p[2] + "-" + String(p[1]).padStart(2, '0') + "-" + String(p[0]).padStart(2, '0');
      }
    }
  }
  return str;
}

// 5. DOWNLOAD DATA (UPDATED FOR FULL & ROBUST FILTERING)
function getRawDataForExcel(tglAwal, tglAkhir, filterNama, filterJenis, role, user, filterKelas, filterKategori) {
  const s = getSheet('Laporan');
  if(s.getLastRow() < 2) return [];

  const sheetSantri = getSheet('Data_Santri');
  let mapKelas = {}; 
  let mapMusyrif = {};
  if (sheetSantri.getLastRow() >= 2) {
    let dataSantri = sheetSantri.getRange(2, 1, sheetSantri.getLastRow()-1, 3).getValues();
    dataSantri.forEach(r => { 
      var sn = String(r[0] || "").trim().toUpperCase();
      if (sn) {
        mapKelas[sn] = String(r[1] || "-").trim();
        mapMusyrif[sn] = String(r[2] || "").trim();
      }
    });
  }

  let rawData = s.getDataRange().getDisplayValues();
  rawData.shift(); 

  let fNama = (!filterNama || filterNama === "" || filterNama === "ALL") ? "ALL" : String(filterNama).trim().toUpperCase();
  let fJenis = (!filterJenis || filterJenis === "" || filterJenis === "ALL") ? "ALL" : String(filterJenis).trim().toUpperCase();
  let fKelas = (!filterKelas || filterKelas === "" || filterKelas === "ALL") ? "ALL" : String(filterKelas).trim().toUpperCase();
  let fKategori = (!filterKategori || filterKategori === "" || filterKategori === "ALL") ? "ALL" : String(filterKategori).trim().toLowerCase();
  let rRole = String(role || "").trim().toLowerCase();
  let isAdmin = (rRole === 'admin');
  let rUser = String(user || "").trim();
  let userAliases = !isAdmin ? getMusyrifAliases(rUser) : [];

  let filtered = rawData.filter(r => {
    var tglISO = convertToISODate(r[0]);
    
    let dateMatch = (!tglAwal || tglISO >= tglAwal) && (!tglAkhir || tglISO <= tglAkhir);
    let rowNama = String(r[3] || "").trim().toUpperCase();
    let rowJenis = String(r[4] || "").trim().toUpperCase();
    let rowSurat = String(r[5] || "").trim().toLowerCase();
    let rowPelapor = String(r[12] || "").trim();
    
    let nameMatch = (fNama === "ALL") ? true : (rowNama === fNama);
    let jenisMatch = (fJenis === "ALL") ? true : (rowJenis === fJenis);
    let santriMusyrif = mapMusyrif[rowNama] || '';
    let userMatch = isAdmin ? true : (matchUserOrMusyrif(santriMusyrif, userAliases) || matchUserOrMusyrif(rowPelapor, userAliases));
    
    let santriKelas = mapKelas[rowNama] || '-';
    let kelasMatch = (fKelas === "ALL") ? true : (santriKelas.toUpperCase() === fKelas);

    let katMatch = true;
    if (fKategori !== "all") {
      let isIqra = (rowSurat.includes('iqra') || rowSurat.includes('tahap') || rowSurat.includes('jilid'));
      let isJuz = rowSurat.includes('juz');
      if (fKategori === "tahap") katMatch = isIqra;
      else if (fKategori === "juz") katMatch = isJuz && !isIqra;
      else if (fKategori === "surat") katMatch = !isIqra && !isJuz;
    }

    return dateMatch && nameMatch && jenisMatch && userMatch && kelasMatch && katMatch;
  });

  return filtered.map(r => {
    let namaUP = String(r[3] || "").trim().toUpperCase();
    return {
      tanggal: r[0], 
      mulai: r[1],
      selesai: r[2],
      nama: r[3], 
      kelas: mapKelas[namaUP] || '-', 
      jenis: r[4], 
      surat: r[5], 
      halAwal: r[6],
      halAkhir: r[7],
      ayatAwal: r[8],
      ayatAkhir: r[9],
      nilai: r[10], 
      catatan: r[11], 
      pelapor: r[12]
    };
  });
}

// 6. CRUD SANTRI (SIMPAN BIODATA, PASSWORD, TARGET, NO WA WALI, DAN USERNAME WALI)
function tambahSantri(n, c, m, passWali, target, noWa, usernameWali){ 
  var sheet = getSheet('Data_Santri');
  if(sheet.getMaxColumns() < 7) sheet.insertColumnsAfter(sheet.getMaxColumns(), 7 - sheet.getMaxColumns());
  var cleanNoWa = noWa ? formatNomorWA(noWa) : '';
  var cleanUserWali = usernameWali ? String(usernameWali).trim() : '';
  sheet.appendRow([n, c, m, passWali || '', target || '', cleanNoWa, cleanUserWali]); 
  return JSON.stringify({
    status: "SUKSES",
    message: "Santri & Data Wali Berhasil Ditambahkan",
    row: sheet.getLastRow()
  }); 
}

function updateSantri(r, n, c, m, passWali, target, noWa, usernameWali){ 
  try{ 
    var sheet = getSheet('Data_Santri');
    if(sheet.getMaxColumns() < 7) sheet.insertColumnsAfter(sheet.getMaxColumns(), 7 - sheet.getMaxColumns());
    var cleanNoWa = (noWa !== undefined && noWa !== null) ? formatNomorWA(noWa) : '';
    var cleanUserWali = (usernameWali !== undefined && usernameWali !== null) ? String(usernameWali).trim() : '';
    
    if(passWali && String(passWali).trim() !== "") {
        sheet.getRange(r, 1, 1, 7).setValues([[n, c, m, passWali, target || '', cleanNoWa, cleanUserWali]]);
    } else {
        sheet.getRange(r, 1, 1, 3).setValues([[n, c, m]]);
        sheet.getRange(r, 5, 1, 3).setValues([[target || '', cleanNoWa, cleanUserWali]]);
    }
    return "SUKSES: Data Santri Diupdate"; 
  } catch(e){ return "ERROR: " + e; } 
}

function editSantri(r, n, c, m, passWali, target, noWa, usernameWali){
  return updateSantri(r, n, c, m, passWali, target, noWa, usernameWali);
}

function hapusSantri(row) {
  try {
    var r = parseInt(row, 10);
    if (!r) return "ERROR: Row tidak valid";
    var sheet = getSheet('Data_Santri');
    sheet.deleteRow(r);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusData(s,r){ try{ getSheet(s).deleteRow(r); return "SUKSES"; }catch(e){return "ERROR";} }

function getAppSettings() {
  const s = getSheet('Pengaturan');
  // Initialize keys if they don't exist
  const keys = ["App Name", "App Logo Data", "App Address", "App Contact", "App Website", "App Kop Gambar"];
  const defaults = ["Mutaba'ah Digital", "", "-", "-", "-", ""];
  
  let settings = {};
  keys.forEach((key, i) => {
    let found = false;
    let lastRow = s.getLastRow();
    if (lastRow >= 2) {
      let data = s.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let r = 0; r < data.length; r++) {
        if (data[r][0] === key) {
          settings[key] = data[r][1];
          found = true;
          break;
        }
      }
    }
    
    if (!found) {
      s.appendRow([key, defaults[i]]);
      settings[key] = defaults[i];
    }
  });

  return { 
    name: settings["App Name"], 
    logoUrl: settings["App Logo Data"],
    address: settings["App Address"],
    contact: settings["App Contact"],
    website: settings["App Website"],
    kopGambar: settings["App Kop Gambar"]
  };
}

function saveAppSettings(name, logoData, address, contact, website, kopGambar) {
  const s = getSheet('Pengaturan');
  const data = s.getDataRange().getValues();
  
  const updates = {
    "App Name": name,
    "App Address": address,
    "App Contact": contact,
    "App Website": website
  };
  
  // Handle Logo: If DELETE, clear it. If new data, set it. If null/empty/undefined, ignore (keep old)
  if (logoData === "DELETE") updates["App Logo Data"] = ""; 
  else if (logoData && logoData !== "") updates["App Logo Data"] = logoData;
  
  // Handle Kop: Same logic
  if (kopGambar === "DELETE") updates["App Kop Gambar"] = "";
  else if (kopGambar && kopGambar !== "") updates["App Kop Gambar"] = kopGambar;
  
  for (let key in updates) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
       if (data[i][0] === key) {
         s.getRange(i + 1, 2).setValue(updates[key]);
         found = true;
         break;
       }
    }
    if (!found) s.appendRow([key, updates[key]]);
  }
  
  return "Pengaturan Tersimpan";
}
function tambahUser(u,p,n,r){ getSheet('Data_User').appendRow([u,p,n,r]); return "SUKSES"; }

// 7b. GET LIST USER (UNTUK ADMIN, WITH SERVER-SIDE SEARCH)
function getListUser(searchQuery) {
  const s = getSheet('Data_User');
  if(s.getLastRow() < 2) return [];
  // Kolom: Username, Password, Nama, Role
  const d = s.getRange(2, 1, s.getLastRow()-1, 4).getValues();
  
  let allData = d.map((row, i) => ({
    username: row[0],
    nama: row[2],
    role: row[3],
    row: i + 2
  })).filter(x => x.username);

  // Server-side search filter
  if (searchQuery && searchQuery.trim() !== "") {
    let q = searchQuery.toLowerCase();
    allData = allData.filter(x => 
      String(x.username).toLowerCase().includes(q) || 
      String(x.nama).toLowerCase().includes(q) ||
      String(x.role).toLowerCase().includes(q)
    );
  }

  return allData;
}

// 7c. UPDATE USER
function updateUser(rowNum, data) {
  try {
    var sheet = getSheet('Data_User');
    if(data.password && data.password !== "") {
      // Update dengan password baru
      sheet.getRange(rowNum, 1, 1, 4).setValues([[data.username, data.password, data.nama, data.role]]);
    } else {
      // Update tanpa password (username, nama, role saja - skip password column)
      sheet.getRange(rowNum, 1).setValue(data.username);
      sheet.getRange(rowNum, 3).setValue(data.nama);
      sheet.getRange(rowNum, 4).setValue(data.role);
    }
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// 7d. HAPUS USER
function hapusUser(rowNum) {
  try {
    var sheet = getSheet('Data_User');
    sheet.deleteRow(rowNum);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function getSheet(name) {
  let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) { 
      s = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name); 
      if(name=='Pengaturan') s.appendRow(['Key','Value']); 
      if(name=='Laporan') s.appendRow(['Tanggal', 'Mulai', 'Selesai', 'Nama', 'Jenis', 'Surat', 'Hal Awal', 'Hal Akhir', 'Ayat Awal', 'Ayat Akhir', 'Nilai', 'Catatan', 'Musyrif']);
      if(name=='Data_Santri') s.appendRow(['Nama Santri', 'Kelas', 'Musyrif', 'Password', 'Target', 'No WA Wali', 'Username Wali']); 
      if(name=='Data_Surat') {
          s.appendRow(['Nama Surat']);
          DAFTAR_114_SURAT.forEach(function(srt, idx) {
              s.appendRow([(idx + 1) + ". " + srt]);
          });
      }
      if(name=='Absensi') s.appendRow(['Tanggal', 'Waktu', 'Nama Santri', 'Status', 'Musyrif']);
      if(name=='Data_Raport') s.appendRow(["Timestamp", "ID", "Nama", "Semester", "Tahun", "Z_P", "Z_K", "Z_C", "M_P", "M_K", "M_C", "T_P", "T_K", "T_C", "Ket_Tambahan", "Grade", "Lanjut", "Sakit", "Izin", "Alpa", "Diberikan_Di", "Tanggal_Raport", "Pembimbing"]);
      if(name=='Kategori_Absensi') {
          s.appendRow(['Kategori Kegiatan', 'Waktu Default']);
          s.appendRow(['Tahfidz - Subuh', "04:30 - 06:30"]);
          s.appendRow(['Tahfidz - Isya', "19:30 - 21:30"]);
          s.appendRow(['Tahfidz - Muroja\'ah', "08:00 - 11:00"]);
          s.appendRow(['Sholat Berjamaah', "04:00 - 20:30"]);
          s.appendRow(['Kajian / Halaqah', "18:00 - 19:30"]);
          s.appendRow(['Apel Harian', "07:30 - 08:00"]);
          s.appendRow(['Kebersihan / Ro\'an', "15:30 - 17:00"]);
          s.appendRow(['Kegiatan Khusus Pondok', "13:30 - 17:00"]);
          s.appendRow(['Kegiatan / Jadwal Libur', "- (Bebas)"]);
      }
      if(name=='Target_Hafalan') {
          s.appendRow(['Tipe', 'Target Untuk', 'Kategori Target', 'Detail Target', 'Batas Waktu', 'Keterangan']);
          s.appendRow(['Kelas', '7A', 'Juz', 'Juz 30 (An-Naba s/d An-Nas)', 'Semester Ganjil', 'Target standar kelas 7']);
          s.appendRow(['Kelas', '8A', 'Juz', 'Juz 29 (Al-Mulk s/d Al-Mursalat)', 'Semester Ganjil', 'Target standar kelas 8']);
      }
      if(name=='Master_Kelas') {
          s.appendRow(['Nama Kelas', 'Jenjang / Tingkat', 'Target Hafalan Default', 'Keterangan']);
          s.appendRow(['7A', 'SMP / MTs', 'Juz 30 (An-Naba s/d An-Nas)', 'Kelas Tahfidz Dasar']);
          s.appendRow(['7B', 'SMP / MTs', 'Juz 30 (An-Naba s/d An-Nas)', 'Kelas Tahfidz Dasar']);
          s.appendRow(['8A', 'SMP / MTs', 'Juz 29 (Al-Mulk s/d Al-Mursalat)', 'Kelas Tahfidz Menengah']);
          s.appendRow(['8B', 'SMP / MTs', 'Juz 29 (Al-Mulk s/d Al-Mursalat)', 'Kelas Tahfidz Menengah']);
          s.appendRow(['9A', 'SMP / MTs', 'Tahap 1 (Jilid 1 & 2)', 'Kelas Tahfidz Lanjutan']);
          s.appendRow(['9B', 'SMP / MTs', 'Surah Al-Baqarah Ayat 1 - 50', 'Kelas Tahfidz Lanjutan']);
      }
      if(name=='Master_Quran') {
          s.appendRow(['No Halaman', 'Rentang Ayat']);
      }
  }
  return s;
}

// --- 8. FITUR DOWNLOAD KHUSUS APK (SERVER SIDE) ---
function generateExcelUrl(dStart, dEnd, fNama, fJenis, role, user, fKelas, mode, fKategori, appName) {
  try {
    // 1. Ambil Data Mentah (Sudah difilter lengkap dengan Kategori & case insensitive)
    var data = getRawDataForExcel(dStart, dEnd, fNama, fJenis, role, user, fKelas || "", fKategori || ""); 
    
    if (mode === 'rekap') {
       data = getAggregatedDataServer(data);
    }
    
    if (!data || data.length === 0) return "KOSONG";

    // 2. Buat Spreadsheet Sementara di Drive
    var namaLembaga = appName || "Mutaba'ah Digital";
    var fileName = "Laporan_" + (mode === 'rekap' ? "Rekap_" : "Detail_") + dStart + "_sd_" + dEnd + "_" + Date.now(); 
    var tempSS = SpreadsheetApp.create(fileName);
    var sheet = tempSS.getActiveSheet();
    sheet.setName("Laporan");
    
    // 3. Siapkan Header Title & Tabel yang Rapi
    var infoSantri = (!fNama || fNama === "ALL") ? "Semua Santri" : fNama;
    var infoKelas = (!fKelas || fKelas === "ALL") ? "Semua Kelas" : fKelas;
    var infoJenis = (!fJenis || fJenis === "ALL") ? "Semua Setoran" : fJenis;
    var infoKategori = (!fKategori || fKategori === "ALL") ? "Semua Kategori" : fKategori;
    var periodStr = dStart + " s/d " + dEnd;
    
    // Tulis Judul Laporan (Baris 1 - 3)
    sheet.getRange("A1").setValue(namaLembaga.toUpperCase()).setFontWeight("bold").setFontSize(15).setFontColor("#1a7d56");
    sheet.getRange("A2").setValue("Laporan Hafalan & Muroja'ah (" + (mode === 'rekap' ? "Ringkasan" : "Detail") + ")").setFontWeight("bold").setFontSize(12);
    sheet.getRange("A3").setValue("Periode: " + periodStr + " | Kelas: " + infoKelas + " | Santri: " + infoSantri + " | Kategori: " + infoKategori + " | Jenis: " + infoJenis).setFontSize(10).setFontColor("#475569");

    var headers = [["TANGGAL", "MULAI", "SELESAI", "NAMA SANTRI", "KELAS", "JENIS", "SURAT / JUZ", "HAL AWAL", "HAL AKHIR", "AYAT AWAL", "AYAT AKHIR", "NILAI", "CATATAN", "MUSYRIF"]];
    
    var rows = data.map(function(r){
      return [
        r.tanggal || "-", 
        r.mulai || "-", 
        r.selesai || "-", 
        r.nama || "-", 
        r.kelas || "-", 
        r.jenis || "-", 
        r.surat || "-", 
        r.halAwal || "-", 
        r.halAkhir || "-", 
        r.ayatAwal || "-", 
        r.ayatAkhir || "-", 
        r.nilai || "-", 
        r.catatan || "-", 
        r.pelapor || "-"
      ];
    });
    
    // 4. Tulis Header Tabel di baris 5
    var headerRange = sheet.getRange(5, 1, 1, 14);
    headerRange.setValues(headers)
               .setFontWeight("bold")
               .setBackground("#1a7d56")
               .setFontColor("white")
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle");
               
    // 5. Tulis Data dan Styling Rapi
    if (rows.length > 0) {
      var dataRange = sheet.getRange(6, 1, rows.length, 14);
      dataRange.setValues(rows)
               .setVerticalAlignment("middle")
               .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
               
      // Alignment per kolom agar rapi dan profesional
      sheet.getRange(6, 1, rows.length, 3).setHorizontalAlignment("center"); // Tgl, Mulai, Selesai
      sheet.getRange(6, 4, rows.length, 1).setHorizontalAlignment("left").setFontWeight("bold"); // Nama
      sheet.getRange(6, 5, rows.length, 2).setHorizontalAlignment("center"); // Kelas, Jenis
      sheet.getRange(6, 7, rows.length, 1).setHorizontalAlignment("left"); // Surat/Juz
      sheet.getRange(6, 8, rows.length, 5).setHorizontalAlignment("center"); // Hal Awal-Akhir, Ayat Awal-Akhir, Nilai
      sheet.getRange(6, 13, rows.length, 1).setHorizontalAlignment("left").setWrap(true); // Catatan wajib wrap agar tidak berantakan
      sheet.getRange(6, 14, rows.length, 1).setHorizontalAlignment("center"); // Musyrif
      
      // Beri warna latar untuk kolom Nilai (A = Hijau muda, B = Kuning muda, C = Merah muda)
      for(var i = 0; i < rows.length; i++) {
         var valNilai = rows[i][11];
         var cellNilai = sheet.getRange(6 + i, 12);
         if(valNilai === 'A') cellNilai.setBackground("#d1fae5").setFontWeight("bold");
         else if(valNilai === 'B') cellNilai.setBackground("#fef3c7").setFontWeight("bold");
         else if(valNilai === 'C') cellNilai.setBackground("#ffe4e6").setFontWeight("bold");
      }
    }
    
    // Atur Lebar Kolom Secara Presisi agar Tidak Berantakan atau Terpotong
    var colWidths = [95, 60, 60, 200, 80, 70, 160, 75, 75, 75, 75, 55, 260, 130];
    for (var col = 0; col < colWidths.length; col++) {
      sheet.setColumnWidth(col + 1, colWidths[col]);
    }
    
    // 6. Atur Izin Akses (PENTING BIAR APK BISA DOWNLOAD)
    var fileId = tempSS.getId();
    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 7. Return Link Download (Format XLSX)
    var url = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
    return url;
    
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

// --- 9. ABSENSI TAHFIDZ & KEGIATAN FUNCTIONS ---
// Kelola Kategori Absensi
function getKategoriAbsensi() {
  try {
    var s = getSheet('Kategori_Absensi');
    var lastRow = s.getLastRow();
    if (lastRow < 2) return [];
    var data = s.getRange(2, 1, lastRow - 1, 2).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] !== "") {
        result.push({
          row: i + 2,
          kategori: String(data[i][0]).trim(),
          waktu: data[i][1] ? String(data[i][1]).trim() : "-"
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

function tambahKategoriAbsensi(kategori, waktu) {
  try {
    if (!kategori) return "Kategori tidak boleh kosong";
    var s = getSheet('Kategori_Absensi');
    s.appendRow([kategori, waktu || ""]);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function updateKategoriAbsensi(rowNum, kategori, waktu) {
  try {
    var s = getSheet('Kategori_Absensi');
    s.getRange(rowNum, 1, 1, 2).setValues([[kategori, waktu || ""]]);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusKategoriAbsensi(rowNum) {
  try {
    var s = getSheet('Kategori_Absensi');
    s.deleteRow(rowNum);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusKategoriAbsensiSeleksi(rows) {
  try {
    var s = getSheet('Kategori_Absensi');
    // Sort descending agar indeks tidak bergeser saat dihapus
    rows.sort(function(a, b){ return b - a; });
    var count = 0;
    rows.forEach(function(r) {
       s.deleteRow(r);
       count++;
    });
    return "SUKSES: " + count + " kategori dihapus";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// --- FITUR TARGET HAFALAN ---
function getTargetHafalan() {
  try {
    var s = getSheet('Target_Hafalan');
    if (s.getLastRow() < 2) return [];
    var data = s.getDataRange().getDisplayValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== "" || String(data[i][1]).trim() !== "") {
        list.push({
          row: i + 1,
          tipe: String(data[i][0]).trim() || 'Kelas',
          targetUntuk: String(data[i][1]).trim(),
          kategori: String(data[i][2]).trim() || 'Juz',
          detail: String(data[i][3]).trim(),
          batasWaktu: String(data[i][4]).trim() || '-',
          keterangan: String(data[i][5]).trim() || '-'
        });
      }
    }
    return list;
  } catch(e) {
    return [];
  }
}

function simpanTargetHafalan(data) {
  try {
    var s = getSheet('Target_Hafalan');
    var row = Number(data.row || 0);
    var rowVal = [
      data.tipe || 'Kelas',
      data.targetUntuk || '',
      data.kategori || 'Juz',
      data.detail || '',
      data.batasWaktu || '-',
      data.keterangan || '-'
    ];
    if (row > 1) {
      s.getRange(row, 1, 1, 6).setValues([rowVal]);
    } else {
      s.appendRow(rowVal);
    }
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusTargetHafalan(rowNum) {
  try {
    var s = getSheet('Target_Hafalan');
    s.deleteRow(rowNum);
    return "SUKSES";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusTargetHafalanSeleksi(rows) {
  try {
    var s = getSheet('Target_Hafalan');
    rows.sort(function(a, b){ return b - a; });
    var count = 0;
    rows.forEach(function(r) {
       s.deleteRow(r);
       count++;
    });
    return "SUKSES: " + count + " target dihapus";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// --- FITUR MASTER KELAS (CRUD & TARGET) ---
function getMasterKelas() {
  try {
    var s = getSheet('Master_Kelas');
    if (s.getLastRow() < 2) return [];
    var data = s.getDataRange().getDisplayValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== "") {
        result.push({
          row: i + 1,
          nama: String(data[i][0]).trim(),
          jenjang: data[i][1] ? String(data[i][1]).trim() : "-",
          target: data[i][2] ? String(data[i][2]).trim() : "Juz 30 (An-Naba s/d An-Nas)",
          keterangan: data[i][3] ? String(data[i][3]).trim() : "-"
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

function simpanMasterKelas(data) {
  try {
    if (!data.nama) return "Nama Kelas tidak boleh kosong";
    var s = getSheet('Master_Kelas');
    var r = parseInt(data.row || 0, 10);
    var rowVal = [data.nama, data.jenjang || "", data.target || "Juz 30 (An-Naba s/d An-Nas)", data.keterangan || ""];
    if (r > 0) {
      s.getRange(r, 1, 1, 4).setValues([rowVal]);
      return "SUKSES: Master Kelas & Target diperbarui";
    } else {
      s.appendRow(rowVal);
      return JSON.stringify({
        status: "SUKSES",
        message: "Master Kelas & Target ditambahkan",
        row: s.getLastRow()
      });
    }
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusMasterKelas(rowNum) {
  try {
    var s = getSheet('Master_Kelas');
    s.deleteRow(rowNum);
    return "SUKSES: Master Kelas dihapus";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

function hapusMasterKelasSeleksi(rowArray) {
  try {
    var s = getSheet('Master_Kelas');
    var rows = rowArray.map(function(item){ return parseInt(item, 10); }).sort(function(a, b){ return b - a; });
    var count = 0;
    rows.forEach(function(r) {
       s.deleteRow(r);
       count++;
    });
    return "SUKSES: " + count + " Master Kelas dihapus";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// Load list santri plus existing status untuk tanggal/waktu tertentu
function loadSantriAbsensi(tanggal, waktu, user, role) {
  try {
    const sSantri = getSheet('Data_Santri');
    const sAbsen = getSheet('Absensi');
    
    // 1. Get All Santri
    if (sSantri.getLastRow() < 2) return [];
    const santriData = sSantri.getRange(2, 1, sSantri.getLastRow()-1, 3).getValues();
    let listSantri = santriData.map(r => ({ nama: String(r[0] || '').trim(), kelas: String(r[1] || '').trim(), musyrif: String(r[2] || '').trim(), status: '', catatan: '' })).filter(x => x.nama);
    
    var rLower = String(role || '').toLowerCase();
    if (rLower !== 'admin') {
      var uAliases = getMusyrifAliases(user);
      listSantri = listSantri.filter(x => matchUserOrMusyrif(x.musyrif, uAliases));
    }
    
    // 2. Get Existing Absensi for the date/time
    if (sAbsen.getLastRow() < 2) return listSantri;
    
    // Format tanggal ke yyyy-MM-dd untuk pencocokan robust
    const targetDateISO = tanggal; // Input tanggal is yyyy-MM-dd
    
    const absenData = sAbsen.getDataRange().getValues();
    absenData.shift(); // remove header
    
    const mappedStatus = {};
    const mappedCatatan = {};
    absenData.forEach(r => {
      // r[0]: Tgl, r[1]: Waktu, r[2]: Nama, r[3]: Status
      let rowDateISO = "";
      if (r[0] instanceof Date) {
        rowDateISO = Utilities.formatDate(r[0], "Asia/Jakarta", "yyyy-MM-dd");
      } else {
        const tStr = String(r[0]);
        const p = tStr.split('/');
        if (p.length === 3) {
          rowDateISO = p[2] + "-" + p[1].padStart(2, '0') + "-" + p[0].padStart(2, '0');
        } else if (tStr.includes('-')) {
          rowDateISO = tStr.split('T')[0];
        }
      }

      if (rowDateISO === targetDateISO && String(r[1]).trim() === String(waktu).trim()) {
        mappedStatus[r[2]] = r[3];
        mappedCatatan[r[2]] = r[5] || "";
      }
    });
    
    // 3. Merge
    listSantri.forEach(s => {
      if (mappedStatus[s.nama]) s.status = mappedStatus[s.nama];
      if (mappedCatatan[s.nama]) s.catatan = mappedCatatan[s.nama];
    });
    
    return listSantri;
  } catch (e) {
    return { error: e.toString() };
  }
}

// Simpan absensi (overwrite existing for date/time/nama)
function simpanAbsensi(listData) {
  try {
    if (!listData || listData.length === 0) return "DATA KOSONG";
    const sheet = getSheet('Absensi');
    
    // Ambil info dari item pertama
    const sample = listData[0];
    const parts = sample.tanggal.split('-');
    const formattedDate = parts[2] + "/" + parts[1] + "/" + parts[0];
    const waktu = sample.waktu;
    
    // 1. Bersihkan data lama untuk tanggal & waktu ini (biar tidak dobel jika diinput ulang)
    if (sheet.getLastRow() >= 2) {
      const data = sheet.getDataRange().getValues();
      const targetDateISO = sample.tanggal; // yyyy-MM-dd
      
      for (var i = data.length - 1; i >= 1; i--) {
        let rowDateISO = "";
        if (data[i][0] instanceof Date) {
          rowDateISO = Utilities.formatDate(data[i][0], "Asia/Jakarta", "yyyy-MM-dd");
        } else {
          const tStr = String(data[i][0]);
          const p = tStr.split('/');
          if (p.length === 3) {
            rowDateISO = p[2] + "-" + p[1].padStart(2, '0') + "-" + p[0].padStart(2, '0');
          } else if (tStr.includes('-')) {
            rowDateISO = tStr.split('T')[0];
          }
        }
        
        if (rowDateISO === targetDateISO && String(data[i][1]).trim() === String(waktu).trim()) {
          sheet.deleteRow(i + 1);
        }
      }
    }
    
    // 2. Insert Batch Baru menggunakan setValues (Jauh lebih cepat dari appendRow loop)
    const rowsToAppend = listData.map(item => [
      formattedDate,
      item.waktu,
      item.nama,
      item.status,
      item.pelapor,
      item.catatan || ""
    ]);
    
    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    }
    
    return "SUKSES";
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

// Simpan absensi batch (Old version - kept for compatibility if needed)
function simpanAbsensiBatch(tanggal, waktu, absensiJson, musyrif) {
  try {
    var absensiArr = JSON.parse(absensiJson);
    var sheet = getSheet('Absensi');
    
    // Format tanggal dari ISO ke dd/MM/yyyy
    var parts = tanggal.split('-');
    var formattedDate = parts[2] + "/" + parts[1] + "/" + parts[0];
    
    // Append each santri absensi
    absensiArr.forEach(function(item) {
      sheet.appendRow([
        formattedDate, // Tanggal
        waktu,         // Waktu (Subuh/Isya)
        item.nama,     // Nama Santri
        item.status,   // Status (H/I/A)
        musyrif        // Musyrif
      ]);
    });
    
    return "SUKSES: " + absensiArr.length + " santri diabsen";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// Get rekap absensi per santri
function getRekapAbsensi(tglAwal, tglAkhir, user, role) {
  try {
    const s = getSheet('Absensi');
    if (s.getLastRow() < 2) return [];

    const sheetSantri = getSheet('Data_Santri');
    let mapSantriMusyrif = {};
    if (sheetSantri.getLastRow() >= 2) {
      let dataSantri = sheetSantri.getRange(2, 1, sheetSantri.getLastRow()-1, 3).getValues();
      dataSantri.forEach(r => {
        var sName = String(r[0] || '').trim();
        if (sName) mapSantriMusyrif[sName] = String(r[2] || '').trim();
      });
    }

    const displayData = s.getDataRange().getValues();
    displayData.shift(); // Remove header
    const lastRow = s.getLastRow();
    displayData.reverse();

    let filtered = [];
    const roleLower = String(role || '').toLowerCase();
    const isAdmin = (roleLower === 'admin' || roleLower === 'pimpinan');
    const userTrim = String(user || '').trim();
    const userAliases = !isAdmin ? getMusyrifAliases(userTrim) : [];

    displayData.forEach((r, i) => {
      var rawDate = r[0];
      var tglISO = "";

      // 1. Parse Date to ISO yyyy-MM-dd
      if (rawDate instanceof Date) {
        tglISO = Utilities.formatDate(rawDate, "Asia/Jakarta", "yyyy-MM-dd");
        r[0] = Utilities.formatDate(rawDate, "Asia/Jakarta", "dd/MM/yyyy"); // Format display Consistency
      } else {
        var tglStr = String(rawDate);
        var pecah = tglStr.split('/');
        if (pecah.length === 3) {
          tglISO = pecah[2] + "-" + pecah[1].padStart(2, '0') + "-" + pecah[0].padStart(2, '0');
        } else if (tglStr.includes('-')) {
          tglISO = tglStr.split('T')[0];
        }
      }

      // 2. Matching logic
      let dateMatch = (tglISO >= tglAwal && tglISO <= tglAkhir);
      var sName = String(r[2] || '').trim();
      var santriMusyrif = mapSantriMusyrif[sName] || '';
      var pelaporMusyrif = String(r[4] || '').trim();
      let userMatch = isAdmin ? true : (matchUserOrMusyrif(santriMusyrif, userAliases) || matchUserOrMusyrif(pelaporMusyrif, userAliases));

      if (dateMatch && userMatch) {
        filtered.push({
          row: lastRow - i, // Correct row index for reversed data
          tanggal: r[0],
          waktu: r[1],
          nama: r[2],
          status: r[3],
          pelapor: r[4],
          catatan: r[5] || ""
        });
      }
    });

    return filtered;
  } catch (e) {
    return { error: e.toString() };
  }
}

// Update Absensi Record
function updateAbsensi(row, status) {
  try {
     var sheet = getSheet('Absensi');
     sheet.getRange(row, 4).setValue(status); // Column 4 is Status
     return "SUKSES";
  } catch(e) {
     return "ERROR: " + e.toString();
  }
}

// Hapus Absensi Record
function hapusAbsensi(row) {
  try {
     var sheet = getSheet('Absensi');
     sheet.deleteRow(row);
     return "SUKSES";
  } catch(e) {
     return "ERROR: " + e.toString();
  }
}

// Hapus Absensi Batch (Delete All in range)
function hapusAbsensiBatch(tglAwal, tglAkhir) {
  try {
    var sheet = getSheet('Absensi');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "KOSONG";
    
    var values = sheet.getRange(2, 1, lastRow-1, 1).getDisplayValues();
    var rowsToDelete = [];
    
    for(var i = 0; i < values.length; i++) {
        var tglSheet = values[i][0];
        var pecah = tglSheet.split('/');
        var tglISO = "";
        if(pecah.length === 3) tglISO = pecah[2] + "-" + pecah[1] + "-" + pecah[0];
        else tglISO = tglSheet;
        
        if (tglISO >= tglAwal && tglISO <= tglAkhir) {
            rowsToDelete.push(i + 2);
        }
    }
    
    rowsToDelete.sort((a,b) => b-a);
    rowsToDelete.forEach(function(r){
        sheet.deleteRow(r);
    });
    
    return "SUKSES: " + rowsToDelete.length + " data dihapus";
  } catch(e) {
    return "ERROR: " + e.toString();
  } 
}

// Hapus Absensi Berdasarkan Seleksi (Array Row)
function hapusAbsensiSeleksi(rows) {
  try {
    var sheet = getSheet('Absensi');
    // Sort descending agar indeks tidak bergeser saat dihapus
    rows.sort(function(a, b){return b - a});
    
    var deletedCount = 0;
    rows.forEach(function(r) {
       sheet.deleteRow(r);
       deletedCount++;
    });
    
    return "SUKSES: " + deletedCount + " data terpilih dihapus";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// 10. GANTI PASSWORD
function gantiPassword(username, oldPass, newPass) {
  try {
    var sheet = getSheet('Data_User');
    var data = sheet.getDataRange().getValues();
    
    for(var i=1; i<data.length; i++) {
        // Col 0: Username, Col 1: Password
       if(String(data[i][0]) === String(username) && String(data[i][1]) === String(oldPass)) {
           sheet.getRange(i+1, 2).setValue(newPass); 
           return "SUKSES";
       }
    }
    return "GAGAL: Password lama salah atau user tidak ditemukan.";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// 11. RAPORT SUMMARY
function getRaportSummary(nama, tglAwal, tglAkhir) {
  try {
    // 1. Ambil Kehadiran
    const sheetAbsensi = getSheet('Absensi');
    let hadir = 0, izin = 0, alpa = 0;
    if (sheetAbsensi.getLastRow() >= 2) {
      const dataAbsen = sheetAbsensi.getRange(2, 1, sheetAbsensi.getLastRow()-1, 5).getDisplayValues();
      dataAbsen.forEach(r => {
        let tglISO = "";
        let pecah = r[0].split('/');
        if(pecah.length === 3) tglISO = pecah[2] + "-" + pecah[1].padStart(2, '0') + "-" + pecah[0].padStart(2, '0');
        else tglISO = r[0];

        if (r[2] === nama && tglISO >= tglAwal && tglISO <= tglAkhir) {
          if (r[3] === 'H') hadir++;
          else if (r[3] === 'I' || r[3] === 'S') izin++;
          else if (r[3] === 'A') alpa++;
        }
      });
    }

    // 2. Ambil Hafalan
    const sheetLaporan = getSheet('Laporan');
    let totalSetoran = 0, totalNilai = 0, pointNilai = 0;
    let terakhir = "-";
    let details = [];

    const mapNilai = { 'A+': 95, 'A': 90, 'A-': 85, 'B+': 80, 'B': 75, 'B-': 70, 'C': 60, 'D': 50 };

    if (sheetLaporan.getLastRow() >= 2) {
      const dataLapor = sheetLaporan.getRange(2, 1, sheetLaporan.getLastRow()-1, 13).getDisplayValues();
      dataLapor.forEach(r => {
        let tglISO = "";
        let pecah = r[0].split('/');
        if(pecah.length === 3) tglISO = pecah[2] + "-" + pecah[1].padStart(2, '0') + "-" + pecah[0].padStart(2, '0');
        else tglISO = r[0];

        if (r[3] === nama && tglISO >= tglAwal && tglISO <= tglAkhir) {
          totalSetoran++;
          let n = r[10].trim();
          if (mapNilai[n]) {
            totalNilai += mapNilai[n];
            pointNilai++;
          }
          terakhir = r[5] + " (" + r[6] + ":" + r[8] + ")";
          details.push({
            tgl: r[0],
            jenis: r[4],
            surat: r[5],
            hal: r[6] + "-" + r[7],
            ayat: r[8] + "-" + r[9],
            nilai: r[10]
          });
        }
      });
    }

    let rata = pointNilai > 0 ? Math.round(totalNilai / pointNilai) : 0;

    return {
      hadir: hadir,
      izin: izin,
      alpa: alpa,
      totalSetoran: totalSetoran,
      rataNilai: rata,
      terakhir: terakhir,
      details: details
    };
  } catch (e) { return { error: e.toString() }; }
}

function saveRaportData(obj) {
  try {
    const s = getSheet('Data_Raport');
    const headers = [
      "Timestamp", "ID", "Nama", "Semester", "Tahun", 
      "Z_P", "Z_K", "Z_C", "M_P", "M_K", "M_C", "T_P", "T_K", "T_C", 
      "Ket_Tambahan", "Grade", "Lanjut", "Sakit", "Izin", "Alpa",
      "Diberikan_Di", "Tanggal_Raport", "Pembimbing"
    ];

    if (s.getLastRow() === 0) s.appendRow(headers);

    const rowData = [
      new Date(),
      obj.id || Utilities.getUuid(),
      obj.nama, obj.semester, obj.tahun,
      obj.z_p, obj.z_k, obj.z_c,
      obj.m_p, obj.m_k, obj.m_c,
      obj.t_p, obj.t_k, obj.t_c,
      obj.ket_tambahan, obj.grade, obj.lanjut, 
      obj.sakit, obj.izin, obj.alpa,
      obj.diberikan_di, obj.tanggal_raport, obj.pembimbing
    ];

    if (obj.id) {
      // Update existing
      const data = s.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === obj.id) {
          s.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          return "SUKSES UPDATE";
        }
      }
    }

    s.appendRow(rowData);
    return "SUKSES SIMPAN";
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

function getRaportHistory(namaSantri) {
  try {
    const s = getSheet('Data_Raport');
    if (s.getLastRow() < 2) return [];
    const data = s.getDataRange().getDisplayValues();
    let result = [];
    var filterNama = namaSantri ? String(namaSantri).trim().toUpperCase() : "ALL";

    for (let i = 1; i < data.length; i++) {
      var rowNama = String(data[i][2]).trim().toUpperCase();
      if (filterNama === "ALL" || rowNama === filterNama) {
        result.push({
          id: data[i][1],
          nama: data[i][2],
          semester: data[i][3],
          tahun: data[i][4],
          tanggal_raport: data[i][20] // Updated index for Tanggal_Raport in new 21-column schema
        });
      }
    }
    return result.reverse();
  } catch (e) { return []; }
}

function getRaportDetail(id) {
  try {
    const s = getSheet('Data_Raport');
    const data = s.getDataRange().getDisplayValues();
    const headers = data[0];
    var searchId = String(id).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === searchId) {
        let obj = {};
        headers.forEach((h, idx) => {
          obj[h.toLowerCase()] = data[i][idx];
        });
        return obj;
      }
    }
    return null;
  } catch (e) { return null; }
}
function deleteRaportData(id) {
  try {
    const s = getSheet('Data_Raport');
    const data = s.getDataRange().getDisplayValues();
    var searchId = String(id).trim();
    for (let i = 1; i < data.length; i++) {
                if (String(data[i][1]).trim() === searchId) {
        s.deleteRow(i + 1);
        return "SUKSES: Data raport berhasil dihapus";
      }
    }
    return "Gagal: ID tidak ditemukan";
  } catch (e) { return "Error: " + e.message; }
}
function getRaportAutoData(nama, semester, tahun) {
  try {
    const sSantri = getSheet('Data_Santri');
    const sAbsen = getSheet('Absensi');
    const sLaporan = getSheet('Laporan');

    // Get santri info
    const santriData = sSantri.getDataRange().getDisplayValues();
    let musyrif = "", kelas = "";
    for (let i = 1; i < santriData.length; i++) {
        if (String(santriData[i][0]).trim().toUpperCase() === String(nama).trim().toUpperCase()) {
            kelas = santriData[i][1];
            musyrif = santriData[i][2];
            break;
        }
    }

    // Calculate semester date range
    let tglAwal, tglAkhir;
    let tahunArr = tahun.split('/'); // e.g. "2025/2026"
    if (semester === 'Ganjil') {
        tglAwal = tahunArr[0] + '-07-01'; // Juli
        tglAkhir = tahunArr[0] + '-12-31'; // Desember
    } else {
        tglAwal = tahunArr[1] + '-01-01'; // Januari
        tglAkhir = tahunArr[1] + '-06-30'; // Juni
    }

    // Get Absensi
    const absenData = sAbsen.getLastRow() >= 2 ? sAbsen.getRange(2, 1, sAbsen.getLastRow() - 1, 5).getDisplayValues() : [];
    let hadir = 0, sakit = 0, izin = 0, alpa = 0;
    
    absenData.forEach(r => {
        let rawDate = r[0];
        let tglISO = "";
        if (rawDate) {
            let tglString = String(rawDate);
            let pecah = tglString.split('/');
            if (pecah.length === 3) {
                tglISO = pecah[2] + "-" + pecah[1].padStart(2, '0') + "-" + pecah[0].padStart(2, '0');
            } else if (tglString.includes('-')) {
                tglISO = tglString;
            }
        }
        
        let rowNama = String(r[2]).trim().toUpperCase();
        let targetNama = String(nama).trim().toUpperCase();

        if (tglISO && rowNama === targetNama && tglISO >= tglAwal && tglISO <= tglAkhir) {
            if (r[3] === 'H') hadir++;
            else if (r[3] === 'S') sakit++;
            else if (r[3] === 'I') izin++;
            else if (r[3] === 'A') alpa++;
        }
    });

    // Get Laporan (Setoran)
    const lapData = sLaporan.getDataRange().getDisplayValues();
    lapData.shift(); // Remove header
    
    let progress = { Z: [], M: [], T: [] };
    let gradeCount = { Z: {A:0, B:0, C:0, D:0}, M: {A:0, B:0, C:0, D:0}, T: {A:0, B:0, C:0, D:0} };
    let allGrades = {A:0, B:0, C:0, D:0};
    let latestEntry = null;

    lapData.forEach(r => {
        let tglParts = r[0].split('/');
        if (tglParts.length !== 3) return;
        let tglISO = tglParts[2] + "-" + tglParts[1].padStart(2, '0') + "-" + tglParts[0].padStart(2, '0');
        
        let rowNama = String(r[3]).trim().toUpperCase();
        let targetNama = String(nama).trim().toUpperCase();
        
        if (rowNama === targetNama && tglISO >= tglAwal && tglISO <= tglAkhir) {
            let cat = String(r[4]).toUpperCase();
            // Map old categories to new
            if (cat === 'TQ') cat = 'Z';
            if (cat === 'PT') cat = 'T';
            
            if (progress.hasOwnProperty(cat)) {
                progress[cat].push(r[5]); // Surat/Juz
                
                // Count grades (A, B, C, D only)
                let nStr = String(r[10]).trim().charAt(0).toUpperCase();
                if (['A','B','C','D'].includes(nStr)) {
                    gradeCount[cat][nStr]++;
                    allGrades[nStr]++;
                }
            }

            // Track latest entry for "Lanjut Ke"
            let tglObj = new Date(tglISO);
            if (!latestEntry || tglObj > latestEntry.date || (tglObj.getTime() === latestEntry.date.getTime() && cat === 'Z')) {
                latestEntry = { date: tglObj, surat: r[5], cat: cat };
            }
        }
    });

    // Helper: Get progress string
    const getProgStr = (arr) => {
        if (arr.length === 0) return "-";
        if (arr.length === 1) return arr[0];
        return arr[0] + " s/d " + arr[arr.length - 1];
    };

    // Helper: Calculate grade by most frequent (nilai terbanyak)
    const getMostFrequentGrade = (countObj) => {
        let max = 0, grade = "-";
        // Priority order: A > B > C > D
        ['A','B','C','D'].forEach(g => {
            if (countObj[g] > max) {
                max = countObj[g];
                grade = g;
            }
        });
        return grade;
    };

    // Smart "Lanjut Ke" Logic
    let saranLanjut = "-";
    if (latestEntry && latestEntry.surat) {
        let text = latestEntry.surat;
        let match = text.match(/(.*?)(\d+)$/);
        if (match) {
            let base = match[1];
            let num = parseInt(match[2]);
            saranLanjut = base + (num + 1);
        } else {
            saranLanjut = text + " (Next)";
        }
    }

    // Overall Grade (most frequent from all)
    let finalGrade = getMostFrequentGrade(allGrades);

    return {
        nama: nama,
        kelas: kelas,
        musyrif: musyrif,
        hadir: hadir,
        sakit: sakit,
        izin: izin,
        alpa: alpa,
        z_p: getMostFrequentGrade(gradeCount.Z),
        z_k: getProgStr(progress.Z),
        z_c: progress.Z.length,
        m_p: getMostFrequentGrade(gradeCount.M),
        m_k: getProgStr(progress.M),
        m_c: progress.M.length,
        t_p: getMostFrequentGrade(gradeCount.T),
        t_k: getProgStr(progress.T),
        t_c: progress.T.length,
        grade: finalGrade,
        lanjut: saranLanjut
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * AMBIL DATA DASHBOARD PROFESIONAL (STATISTIK, CHART, LEADERBOARD, INAKTIF)
 */
function getDashboardData(range, user, role) {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tz = "Asia/Jakarta";
    const isAdmin = String(role).toLowerCase() === 'admin';
    
    // 0. Calculate Date Range
    let startDate = new Date(today);
    if (range === '7days') startDate.setDate(today.getDate() - 6);
    else if (range === '30days') startDate.setDate(today.getDate() - 29);
    
    // Helper to parse "dd/MM/yyyy"
    const parseDate = (str) => {
      if (!str || typeof str !== 'string' || !str.includes('/')) return new Date(0);
      const parts = str.split('/');
      return new Date(parts[2], parts[1] - 1, parts[0]);
    };

    // 1. Fetch RAW Data (Single Trip per Sheet)
    const rawSantri = getSheet('Data_Santri').getDataRange().getValues();
    const rawLaporan = getSheet('Laporan').getDataRange().getValues();
    const rawAbsensi = getSheet('Absensi').getDataRange().getValues();

    // 2. Process Santri Basis
    // Headers: [Nama, Kelas, Musyrif]
    const allSantri = rawSantri.slice(1).map(r => ({ nama: r[0], kelas: r[1], musyrif: r[2] })).filter(x => x.nama);
    const userAliases = !isAdmin ? getMusyrifAliases(user) : [];
    const mySantriList = isAdmin ? allSantri : allSantri.filter(s => matchUserOrMusyrif(s.musyrif, userAliases));
    const mySantriNames = mySantriList.map(s => s.nama);

    // 3. Process Laporan (Stats, Trend, Leaderboard, Inactivity)
    let totalSetoran = 0;
    let quranCount = 0;
    let iqraCount = 0;
    let leaderboardMap = {};
    let lastSetoranMap = {};
    let trendMap = {};

    // Initialize Trend Labels based on range
    let diffDays = Math.round((today - startDate) / (1000 * 60 * 60 * 24)) + 1;
    let trendLabels = [];
    for (let i = diffDays - 1; i >= 0; i--) {
      let d = new Date(today); d.setDate(today.getDate() - i);
      let ds = Utilities.formatDate(d, tz, "dd/MM/yyyy");
      let dl = (range === 'today') ? Utilities.formatDate(d, tz, "HH:mm") : Utilities.formatDate(d, tz, "dd/MM");
      trendLabels.push({ full: ds, short: dl });
      trendMap[ds] = 0;
    }

    const inactiveThreshold = new Date(today);
    inactiveThreshold.setDate(today.getDate() - 2);

    // Process Laporan Records
    // Headers: [Tgl, Start, End, Nama, Jenis, Surat, HalAwal, HalAkhir, AyatAwal, AyatAkhir, Nilai, Catatan, Pelapor]
    rawLaporan.slice(1).forEach(r => {
      const nama = String(r[3] || '').trim();
      const pelapor = String(r[12] || '').trim();
      // Keamanan akses: Musyrif hanya mengakses santri binaannya sendiri
      const isMine = isAdmin ? true : (mySantriNames.indexOf(nama) !== -1);

      if (isMine && nama) {
        let tglObj = (r[0] instanceof Date) ? r[0] : parseDate(String(r[0]));
        tglObj.setHours(0, 0, 0, 0);
        const tglStr = Utilities.formatDate(tglObj, tz, "dd/MM/yyyy");

        // Global track for Inactivity
        if (!lastSetoranMap[nama] || tglObj > lastSetoranMap[nama]) {
          lastSetoranMap[nama] = tglObj;
        }

        // Range Filtered Stats
        if (tglObj >= startDate && tglObj <= today) {
          totalSetoran++;
          if (trendMap.hasOwnProperty(tglStr)) trendMap[tglStr]++;
          
          leaderboardMap[nama] = (leaderboardMap[nama] || 0) + 1;

          const surat = String(r[5]).toLowerCase();
          if (surat.includes("iqra") || surat.includes("jilid") || surat.includes("tahap")) iqraCount++;
          else quranCount++;
        }
      }
    });

    // 4. Process Absensi
    // Headers: [Tgl, ..., Musyrif]
    const totalHadir = rawAbsensi.slice(1).filter(r => {
      const nama = String(r[2] || '').trim();
      const pelapor = String(r[4] || '').trim();
      const isMine = isAdmin ? true : (mySantriNames.indexOf(nama) !== -1);
      if (!isMine) return false;

      let tglObj = (r[0] instanceof Date) ? r[0] : parseDate(String(r[0]));
      tglObj.setHours(0,0,0,0);
      return tglObj >= startDate && tglObj <= today;
    }).length;

    // 5. Finalize Summary
    const leaderboard = Object.keys(leaderboardMap)
      .map(n => ({ nama: n, count: leaderboardMap[n] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    const inactiveSantri = mySantriNames.filter(nama => {
      let lastDate = lastSetoranMap[nama];
      return !lastDate || lastDate < inactiveThreshold;
    });

    return {
      stats: {
        totalSantri: mySantriNames.length,
        totalHadir: totalHadir,
        totalSetoran: totalSetoran
      },
      charts: {
        trend: {
          labels: trendLabels.map(l => l.short),
          values: trendLabels.map(l => trendMap[l.full])
        },
        categories: {
          labels: ["Al-Qur'an", "Tahap (1-6)"],
          values: [quranCount, iqraCount]
        }
      },
      leaderboard: leaderboard,
      inactive: inactiveSantri
    };

  } catch (e) {
    return { error: e.toString() };
  }
}

function getAggregatedDataServer(data) {
  if (!data || data.length === 0) return [];
  var groups = {};
  data.forEach(function(r) {
    var key = r.nama + "|" + r.jenis;
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });
  
  var result = [];
  for (var key in groups) {
    var rows = groups[key];
    rows.sort(function(a, b) {
       var dt1 = a.tanggal.split('/').reverse().join('-') + " " + a.mulai;
       var dt2 = b.tanggal.split('/').reverse().join('-') + " " + b.mulai;
       return dt1.localeCompare(dt2);
    });
    var first = rows[0];
    var last = rows[rows.length - 1];
    var suratRange = first.surat;
    if (first.surat !== last.surat) suratRange = first.surat + " s/d " + last.surat;
    
    result.push({
      tanggal: first.tanggal === last.tanggal ? first.tanggal : first.tanggal + " - " + last.tanggal,
      mulai: first.mulai,
      selesai: last.selesai,
      nama: first.nama,
      kelas: first.kelas,
      jenis: first.jenis,
      surat: suratRange,
      halAwal: first.halAwal,
      halAkhir: last.halAkhir,
      ayatAwal: first.ayatAwal,
      ayatAkhir: last.ayatAkhir,
      nilai: "-",
      catatan: rows.map(function(r){ return r.catatan; }).filter(function(c){ return c && c !== '-'; }).join('; ') || "-",
      pelapor: last.pelapor
    });
  }

  // SORT BY NAME ALPHABETICALLY
  result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

  return result;
}

// =============================================
// === MASTER QURAN: HELPER FUNCTIONS ===
// =============================================

/**
 * Merge overlapping/adjacent integer ranges
 * Input: [[1,5], [3,10], [15,20]]
 * Output: [[1,10], [15,20]]
 */
function mergeRanges(ranges) {
  if (!ranges || ranges.length === 0) return [];
  var sorted = ranges.slice().sort(function(a, b) { return a[0] - b[0]; });
  var merged = [sorted[0].slice()];
  for (var i = 1; i < sorted.length; i++) {
    var last = merged[merged.length - 1];
    if (sorted[i][0] <= last[1] + 1) {
      last[1] = Math.max(last[1], sorted[i][1]);
    } else {
      merged.push(sorted[i].slice());
    }
  }
  return merged;
}

/**
 * Hitung total ayat dari array merged ranges
 */
function hitungTotalAyatDariRanges(mergedRanges) {
  var total = 0;
  for (var i = 0; i < mergedRanges.length; i++) {
    total += mergedRanges[i][1] - mergedRanges[i][0] + 1;
  }
  return total;
}

/**
 * Normalisasi nama surat untuk pencocokan fleksibel (mengakomodasi variasi ejaan & transliterasi)
 */
function normalizeSurat(str) {
  if (!str) return "";
  var s = String(str).toLowerCase().trim();
  // Hapus nomor urut di depan nama surat (misal "1. ", "114. ")
  s = s.replace(/^\d+[\s\.\-_:]*/, "").trim();
  // Hapus prefix umum surat
  s = s.replace(/^(?:q\.?s\.?|sura[th]?|quran sura[th]?)\s*(?:-|_|\s)*\s*/i, "").trim();
  
  // Variasi khusus yang sering berbeda
  if (s.indexOf("mudas") !== -1 || s.indexOf("muddas") !== -1 || s.indexOf("mudat") !== -1 || s.indexOf("muddat") !== -1 || s.indexOf("mudath") !== -1 || s.indexOf("muddath") !== -1) return "almuddatsir";
  if (s.indexOf("muzam") !== -1 || s.indexOf("muzzam") !== -1) return "almuzzammil";
  if (s.indexOf("muthaf") !== -1 || s.indexOf("mutaf") !== -1 || s.indexOf("mutaffif") !== -1) return "almuthaffifin";
  if (s.indexOf("mujadil") !== -1 || s.indexOf("mujadal") !== -1) return "almujadilah";
  if (s.indexOf("lahab") !== -1 || s.indexOf("masad") !== -1) return "almasad";
  if (s.indexOf("inshirah") !== -1 || s.indexOf("insyirah") !== -1 || s.indexOf("syarh") !== -1 || s.indexOf("alam nas") !== -1) return "alinsyirah";
  if (s.indexOf("isra") !== -1 || s.indexOf("bani israil") !== -1) return "alisra";
  if (s.indexOf("kahf") !== -1) return "alkahf";
  if (s.indexOf("ghafir") !== -1 || s.indexOf("mumin") !== -1) return "almumin";
  if (s.indexOf("fushilat") !== -1 || s.indexOf("fussilat") !== -1 || s.indexOf("ha mim") !== -1) return "fushilat";
  if (s.indexOf("imran") !== -1) return "aliimran";
  if (s.indexOf("insan") !== -1 || s.indexOf("dahr") !== -1) return "alinsan";
  if (s.indexOf("zalzalah") !== -1 || s.indexOf("zilzal") !== -1) return "alzalzalah";
  if (s.indexOf("kautsar") !== -1 || s.indexOf("kausar") !== -1) return "alkautsar";
  if (s.indexOf("jatsiyah") !== -1 || s.indexOf("jasiyah") !== -1) return "aljatsiyah";
  if (s.indexOf("takatsur") !== -1 || s.indexOf("takasur") !== -1) return "altakatsur";
  if (s.indexOf("syams") !== -1 || s.indexOf("shams") !== -1) return "asyams";
  if (s.indexOf("bayyinah") !== -1 || s.indexOf("bayinah") !== -1) return "albayyinah";
  if (s.indexOf("buruj") !== -1 || s.indexOf("burooj") !== -1) return "alburuj";
  if (s.indexOf("tariq") !== -1 || s.indexOf("thoriq") !== -1) return "attariq";
  if (s.indexOf("ghasyiyah") !== -1 || s.indexOf("ghasiyah") !== -1) return "alghasyiyah";
  
  // Standardisasi transliterasi huruf
  s = s.replace(/sh/g, "sy"); 
  s = s.replace(/th/g, "t");  
  s = s.replace(/ts/g, "t");  
  s = s.replace(/dz/g, "z");  
  s = s.replace(/dh/g, "z");  
  s = s.replace(/oo/g, "u");   
  s = s.replace(/ee/g, "i");   

  // Hapus semua karakter non-alfanumerik (tanda hubung, spasi, koma, apostrof)
  s = s.replace(/[^a-z0-9]/g, "");
  
  // Hapus vokal atau konsonan ganda berurutan (aa -> a, ii -> i, uu -> u)
  s = s.replace(/([a-z0-9])\1+/g, "$1");
  
  // Hapus akhiran 'h' atau 'i' setelah huruf vokal (cth: baqarah -> baqara, fatihah -> fatiha, kahfi -> kahf)
  s = s.replace(/([aiueo])h$/i, "$1");
  
  return s;
}

function isSuratMatch(str1, str2) {
  if (!str1 || !str2) return false;
  var n1 = normalizeSurat(str1);
  var n2 = normalizeSurat(str2);
  if (!n1 || !n2) return false;
  if (n1 === n2) return true;
  
  // Jika panjang setelah normalisasi minimal 4 huruf, periksa substring (misal: "albaqara" vs "baqara", "annisa" vs "nisa")
  if (n1.length >= 4 && n2.length >= 4) {
    if (n1.indexOf(n2) !== -1 || n2.indexOf(n1) !== -1) return true;
  }
  return false;
}

/**
 * Parse string rentang ayat dari Master_Quran (Mendukung multi-surah per sel dan variasi separator)
 * "Al-Baqara 1-5"  → [{ surat: "Al-Baqara", ayatAwal: 1, ayatAkhir: 5 }]
 * "Al-Ikhlas 1-4, Al-Falaq 1-5, An-Nas 1-6" → 3 object surat
 */
function parseSemuaRentangAyat(str) {
  if (!str) return [];
  var strCloned = String(str).trim();
  if (!strCloned) return [];
  
  // Pisahkan jika ada beberapa surat dalam satu sel (koma, baris baru, titik koma, garis miring)
  var parts = strCloned.split(/(?:,|\n|\r|;|\/|\||\bdan\b|\band\b)/i);
  var results = [];
  
  for (var p = 0; p < parts.length; p++) {
    var part = parts[p].trim();
    if (!part) continue;
    part = part.replace(/^[\(\[]+|[\)\]]+$/g, '').trim();
    
    // Pattern dengan rentang: NamaSurat spasi/titik dua, angkaAwal [ - / s/d / sd / s.d / sampai ] angkaAkhir
    var matchRange = part.match(/^(.+?)(?:\s*(?::|\bayat\b)\s*|\s+)(\d+)\s*(?:-|–|—|\s+(?:s\/d|sd|s\.d|sampai|to|-)\s+)\s*(\d+)$/i);
    if (matchRange && isNaN(part)) {
      results.push({
        surat: matchRange[1].trim(),
        ayatAwal: parseInt(matchRange[2]),
        ayatAkhir: parseInt(matchRange[3])
      });
      continue;
    }
    
    // Pattern single ayat atau rentang simpel
    var matchSimple = part.match(/^(.+?)\s+(\d+)\s*[-–—]\s*(\d+)$/);
    if (matchSimple) {
      results.push({
        surat: matchSimple[1].trim(),
        ayatAwal: parseInt(matchSimple[2]),
        ayatAkhir: parseInt(matchSimple[3])
      });
      continue;
    }
    
    // Fallback single ayat: "Al-Fatihah 7", "Al-Baqarah : 1"
    var matchSingle = part.match(/^(.+?)(?:\s*(?::|\bayat\b)\s*|\s+)(\d+)$/i);
    if (matchSingle && isNaN(part)) {
      results.push({
        surat: matchSingle[1].trim(),
        ayatAwal: parseInt(matchSingle[2]),
        ayatAkhir: parseInt(matchSingle[2])
      });
      continue;
    }
  }
  return results;
}

/**
 * Baca dan parse seluruh data Master_Quran
 * Return: Array of { halaman, surat, ayatAwal, ayatAkhir, raw }
 */
function getMasterQuran() {
  try {
    var s = getSheet('Master_Quran');
    if (s.getLastRow() < 2) return [];
    var data = s.getRange(2, 1, s.getLastRow() - 1, 2).getDisplayValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var halaman = parseInt(data[i][0]) || 0;
      var rentangStr = String(data[i][1]).trim();
      if (!halaman || !rentangStr) continue;
      
      var parsedList = parseSemuaRentangAyat(rentangStr);
      for (var p = 0; p < parsedList.length; p++) {
        var parsed = parsedList[p];
        result.push({
          halaman: halaman,
          surat: parsed.surat,
          ayatAwal: parsed.ayatAwal,
          ayatAkhir: parsed.ayatAkhir,
          raw: rentangStr
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

/**
 * Cari halaman mushaf berdasarkan surat dan rentang ayat (Fuzzy matching)
 * Return: { halAwal: number, halAkhir: number } atau null
 */
function cariHalamanByAyat(surat, ayatAwal, ayatAkhir) {
  try {
    var masterData = getMasterQuran();
    if (!masterData || masterData.length === 0) return null;
    
    ayatAwal = parseInt(ayatAwal) || 0;
    ayatAkhir = parseInt(ayatAkhir) || ayatAwal;
    if (!ayatAwal) return null;
    
    var halAwal = null;
    var halAkhir = null;
    
    for (var i = 0; i < masterData.length; i++) {
      var m = masterData[i];
      if (!isSuratMatch(m.surat, surat)) continue;
      
      // Cek apakah ayat input overlap dengan rentang halaman ini
      if (ayatAwal <= m.ayatAkhir && ayatAkhir >= m.ayatAwal) {
        if (halAwal === null || m.halaman < halAwal) halAwal = m.halaman;
        if (halAkhir === null || m.halaman > halAkhir) halAkhir = m.halaman;
      }
    }
    
    if (halAwal !== null) {
      return { halAwal: halAwal, halAkhir: halAkhir };
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Cek overlap setoran Ziyadah berdasarkan histori di Laporan (Fuzzy matching)
 * Return: { status: 'NEW'|'PARTIAL'|'FULL_OVERLAP', ayatBaru: N, ayatLama: N, detail: string }
 */
function cekOverlapZiyadah(nama, surat, ayatAwal, ayatAkhir, halAwal, halAkhir) {
  try {
    ayatAwal = parseInt(ayatAwal) || 0;
    ayatAkhir = parseInt(ayatAkhir) || 0;
    halAwal = parseInt(halAwal) || 0;
    halAkhir = parseInt(halAkhir) || halAwal;
    
    var shLapor = getSheet('Laporan');
    var namaTrim = String(nama).trim().toLowerCase();
    
    // Jika Ayat kosong (hanya mengisikan rentang Halaman), lakukan cek overlap berdasar Halaman
    if (!ayatAwal || !ayatAkhir) {
      if (halAwal > 0 && halAkhir > 0) {
        if (shLapor.getLastRow() < 2) return { status: 'NEW', ayatBaru: halAkhir - halAwal + 1, ayatLama: 0, unit: 'Halaman', detail: 'Halaman baru' };
        var lapData = shLapor.getRange(2, 4, shLapor.getLastRow() - 1, 7).getDisplayValues();
        var existHal = [];
        for (var i = 0; i < lapData.length; i++) {
          var rNm = String(lapData[i][0]).trim().toLowerCase();
          var rJns = String(lapData[i][1]).trim().toUpperCase();
          var rSrt = String(lapData[i][2]).trim();
          var rHa = parseInt(lapData[i][3]) || 0;
          var rHk = parseInt(lapData[i][4]) || rHa;
          if (rNm === namaTrim && rJns === 'Z' && rHa > 0 && rHk > 0) {
            if (!surat || surat === '-' || rSrt === '-' || isSuratMatch(rSrt, surat)) {
              existHal.push([Math.min(rHa, rHk), Math.max(rHa, rHk)]);
            }
          }
        }
        if (existHal.length === 0) return { status: 'NEW', ayatBaru: halAkhir - halAwal + 1, ayatLama: 0, unit: 'Halaman', detail: 'Semua ' + (halAkhir - halAwal + 1) + ' halaman baru' };
        var mergedHal = mergeRanges(existHal);
        var totBaruH = 0;
        for (var pg = Math.min(halAwal, halAkhir); pg <= Math.max(halAwal, halAkhir); pg++) {
          var sAda = false;
          for (var j = 0; j < mergedHal.length; j++) {
            if (pg >= mergedHal[j][0] && pg <= mergedHal[j][1]) { sAda = true; break; }
          }
          if (!sAda) totBaruH++;
        }
        var totInH = (Math.max(halAwal, halAkhir) - Math.min(halAwal, halAkhir)) + 1;
        var totLamaH = totInH - totBaruH;
        if (totBaruH === 0) {
          return { status: 'FULL_OVERLAP', ayatBaru: 0, ayatLama: totInH, unit: 'Halaman', detail: 'Seluruh halaman ' + halAwal + '-' + halAkhir + ' sudah pernah disetorkan sebelumnya' };
        } else if (totBaruH < totInH) {
          return { status: 'PARTIAL', ayatBaru: totBaruH, ayatLama: totLamaH, unit: 'Halaman', detail: totBaruH + ' halaman baru (' + totLamaH + ' sudah pernah)' };
        } else {
          return { status: 'NEW', ayatBaru: totBaruH, ayatLama: 0, unit: 'Halaman', detail: 'Halaman baru' };
        }
      }
      return { status: 'NEW', ayatBaru: 0, ayatLama: 0, unit: 'Ayat', detail: 'Ayat & Halaman tidak terisi' };
    }
    
    if (shLapor.getLastRow() < 2) {
      return { status: 'NEW', ayatBaru: ayatAkhir - ayatAwal + 1, ayatLama: 0, unit: 'Ayat', detail: 'Belum ada riwayat' };
    }
    
    // Baca: Nama(D), Jenis(E), Surat(F), HalAwal(G), HalAkhir(H), AyatAwal(I), AyatAkhir(J)
    var lapData = shLapor.getRange(2, 4, shLapor.getLastRow() - 1, 7).getDisplayValues();
    
    var existingRanges = [];
    var namaTrim = String(nama).trim().toLowerCase();
    
    for (var i = 0; i < lapData.length; i++) {
      var rowNama = String(lapData[i][0]).trim().toLowerCase();
      var rowJenis = String(lapData[i][1]).trim().toUpperCase();
      var rowSurat = String(lapData[i][2]).trim();
      var rowAyatAwal = parseInt(lapData[i][5]) || 0;
      var rowAyatAkhir = parseInt(lapData[i][6]) || 0;
      
      if (rowNama === namaTrim && rowJenis === 'Z' && isSuratMatch(rowSurat, surat) && rowAyatAwal > 0 && rowAyatAkhir > 0) {
        existingRanges.push([rowAyatAwal, rowAyatAkhir]);
      }
    }
    
    if (existingRanges.length === 0) {
      return { status: 'NEW', ayatBaru: ayatAkhir - ayatAwal + 1, ayatLama: 0, detail: 'Semua ' + (ayatAkhir - ayatAwal + 1) + ' ayat baru' };
    }
    
    // Merge existing ranges
    var merged = mergeRanges(existingRanges);
    
    // Hitung ayat baru yang tidak overlap
    var totalBaru = 0;
    for (var a = ayatAwal; a <= ayatAkhir; a++) {
      var sudahAda = false;
      for (var j = 0; j < merged.length; j++) {
        if (a >= merged[j][0] && a <= merged[j][1]) { sudahAda = true; break; }
      }
      if (!sudahAda) totalBaru++;
    }
    
    var totalInput = ayatAkhir - ayatAwal + 1;
    var totalLama = totalInput - totalBaru;
    
    if (totalBaru === 0) {
      return { 
        status: 'FULL_OVERLAP', 
        ayatBaru: 0, 
        ayatLama: totalInput,
        detail: 'Seluruh ayat ' + ayatAwal + '-' + ayatAkhir + ' sudah pernah disetorkan'
      };
    } else if (totalBaru < totalInput) {
      return { 
        status: 'PARTIAL', 
        ayatBaru: totalBaru, 
        ayatLama: totalLama,
        detail: totalBaru + ' ayat baru dari ' + totalInput + ' ayat (' + totalLama + ' sudah pernah)'
      };
    } else {
      return { status: 'NEW', ayatBaru: totalBaru, ayatLama: 0, detail: 'Semua ' + totalBaru + ' ayat baru' };
    }
  } catch (e) {
    return { status: 'NEW', ayatBaru: (parseInt(ayatAkhir) || 0) - (parseInt(ayatAwal) || 0) + 1, ayatLama: 0, detail: 'Error cek: ' + e.toString() };
  }
}

/**
 * Rekalkulasi progres unik santri dari seluruh histori Ziyadah di Laporan (Fuzzy matching & multi-range)
 * Return: { totalAyatUnik, totalHalamanUnik, suratList, perSurat }
 */
function rekalkulasiProgresUnik(nama) {
  try {
    var shLapor = getSheet('Laporan');
    if (shLapor.getLastRow() < 2) {
      return { totalAyatUnik: 0, totalHalamanUnik: 0, suratList: [], perSurat: {} };
    }
    
    var lapData = shLapor.getRange(2, 4, shLapor.getLastRow() - 1, 7).getDisplayValues();
    var namaTrim = String(nama).trim().toLowerCase();
    
    var perSurat = {};
    var suratSet = {};
    
    for (var i = 0; i < lapData.length; i++) {
      var rowNama = String(lapData[i][0]).trim().toLowerCase();
      var rowJenis = String(lapData[i][1]).trim().toUpperCase();
      var rowSurat = String(lapData[i][2]).trim();
      var rowAyatAwal = parseInt(lapData[i][5]) || 0;
      var rowAyatAkhir = parseInt(lapData[i][6]) || 0;
      
      if (rowNama !== namaTrim) continue;
      
      if (rowSurat && rowSurat !== '-') {
        suratSet[rowSurat] = true;
      }
      
      if (rowJenis === 'Z' && rowSurat && rowAyatAwal > 0 && rowAyatAkhir > 0) {
        if (!perSurat[rowSurat]) perSurat[rowSurat] = [];
        perSurat[rowSurat].push([rowAyatAwal, rowAyatAkhir]);
      }
    }
    
    var totalAyatUnik = 0;
    var mergedPerSurat = {};
    
    for (var surat in perSurat) {
      var merged = mergeRanges(perSurat[surat]);
      mergedPerSurat[surat] = merged;
      totalAyatUnik += hitungTotalAyatDariRanges(merged);
    }
    
    // Hitung halaman unik via Master_Quran
    var totalHalamanUnik = 0;
    try {
      var masterData = getMasterQuran();
      if (masterData.length > 0) {
        var halamanSet = {};
        for (var suratKey in mergedPerSurat) {
          var ranges = mergedPerSurat[suratKey];
          for (var j = 0; j < masterData.length; j++) {
            var m = masterData[j];
            if (!isSuratMatch(m.surat, suratKey)) continue;
            for (var k = 0; k < ranges.length; k++) {
              if (ranges[k][0] <= m.ayatAkhir && ranges[k][1] >= m.ayatAwal) {
                halamanSet[m.halaman] = true;
                break;
              }
            }
          }
        }
        totalHalamanUnik = Object.keys(halamanSet).length;
      }
    } catch (e) {}
    
    return {
      totalAyatUnik: totalAyatUnik,
      totalHalamanUnik: totalHalamanUnik,
      suratList: Object.keys(suratSet),
      perSurat: mergedPerSurat
    };
  } catch (e) {
    return { totalAyatUnik: 0, totalHalamanUnik: 0, suratList: [], perSurat: {} };
  }
}

// --- API BRIDGE UNTUK VERCEL ---
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var funcName = payload.func;
    var args = payload.args || [];
    
    if (typeof this[funcName] !== 'function') {
      throw new Error('Fungsi ' + funcName + ' tidak ditemukan di backend.');
    }
    
    var result = this[funcName].apply(this, args);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// --- END API BRIDGE ---
