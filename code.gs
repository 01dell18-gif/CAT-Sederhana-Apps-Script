/**
 * ============================================================================
 * SISTEM UJIAN ONLINE SUMATIF (ANTI-CURANG) - GOOGLE APPS SCRIPT
 * ============================================================================
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Ujian Online Sumatif')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Catatan Penting: Fungsi Setup Database Sheets Otomatis
function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = [
    { name: 'Pengaturan', headers: ['Mapel', 'Password', 'Timer', 'Jumlah Soal', 'Sekolah', 'Logo'] },
    { name: 'Data-Murid', headers: ['Nama Murid', 'Kelas'] },
    { name: 'Jawaban-Murid', headers: ['Timestamp', 'Nama Murid', 'Kelas', 'Skor Benar', 'Total Soal', 'Nilai Akhir', 'Status Kecurangan'] },
    { name: 'Soal1', headers: ['Soal', 'A', 'B', 'C', 'D', 'E', 'Jawaban'] },
    { name: 'Soal2', headers: ['Soal', 'A', 'B', 'C', 'D', 'E', 'Jawaban'] },
    { name: 'Soal3', headers: ['Soal', 'Jawaban'] }
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight('bold');
    }
  });

  // Isi data dummy jika sheet Pengaturan kosong
  const pengSheet = ss.getSheetByName('Pengaturan');
  if (pengSheet.getLastRow() === 1) {
    pengSheet.appendRow(['Matematika', '12345', '60', '10', 'SMA NEGERI 1 UNGGULAN', 'https://via.placeholder.com/80']);
  }
}

// 2. Mengambil Data Awal (Kelas, Pengaturan) secara Cepat
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const pSheet = ss.getSheetByName('Pengaturan');
  const pData = pSheet.getRange(2, 1, 1, 6).getValues()[0] || [];
  
  const settings = {
    mapel: pData[0] || 'Ujian Online',
    password: String(pData[1] || ''),
    timer: Number(pData[2] || 60),
    jumlahSoal: Number(pData[3] || 10),
    sekolah: pData[4] || 'SEKOLAH KITA',
    logo: pData[5] || ''
  };

  const mSheet = ss.getSheetByName('Data-Murid');
  const mData = mSheet.getLastRow() > 1 ? mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 2).getValues() : [];
  
  const listKelas = [...new Set(mData.map(r => r[1]))].filter(String).sort();

  return { settings, listKelas };
}

// Mengambil Nama Murid Berdasarkan Kelas
function getStudentsByClass(kelas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName('Data-Murid');
  const mData = mSheet.getLastRow() > 1 ? mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 2).getValues() : [];
  
  return mData.filter(r => r[1] === kelas).map(r => r[0]).sort();
}

// 3. Mengambil Soal Ujian (Memastikan Soal1, Soal2, Soal3 Terwakili)
function getExamQuestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pData = ss.getSheetByName('Pengaturan').getRange(2, 1, 1, 4).getValues()[0];
  const targetTotal = Number(pData[3] || 10);

  let poolPG = [];
  let poolPGK = [];
  let poolBS = [];

  // Sheet Soal1: PG (A,B,C,D,E)
  const s1 = ss.getSheetByName('Soal1');
  if (s1 && s1.getLastRow() > 1) {
    const data = s1.getRange(2, 1, s1.getLastRow() - 1, 7).getValues();
    data.forEach(r => {
      if(r[0]) poolPG.push({ type: 'PG', question: String(r[0]), options: { A: r[1], B: r[2], C: r[3], D: r[4], E: r[5] }, answer: String(r[6]).trim() });
    });
  }

  // Sheet Soal2: PG Kompleks (A,B,C,D,E)
  const s2 = ss.getSheetByName('Soal2');
  if (s2 && s2.getLastRow() > 1) {
    const data = s2.getRange(2, 1, s2.getLastRow() - 1, 7).getValues();
    data.forEach(r => {
      if(r[0]) poolPGK.push({ type: 'PGK', question: String(r[0]), options: { A: r[1], B: r[2], C: r[3], D: r[4], E: r[5] }, answer: String(r[6]).trim() });
    });
  }

  // Sheet Soal3: Benar / Salah
  const s3 = ss.getSheetByName('Soal3');
  if (s3 && s3.getLastRow() > 1) {
    const data = s3.getRange(2, 1, s3.getLastRow() - 1, 2).getValues();
    data.forEach(r => {
      if(r[0]) poolBS.push({ type: 'BS', question: String(r[0]), options: { 'Benar': 'Benar', 'Salah': 'Salah' }, answer: String(r[1]).trim() });
    });
  }

  // Pengacakan Fungsi Sederhana (Shuffle)
  const shuffle = (array) => array.sort(() => Math.random() - 0.5);

  shuffle(poolPG);
  shuffle(poolPGK);
  shuffle(poolBS);

  let selected = [];

  // Pastikan minimal ada 1 soal dari masing-masing tipe jika ketersediaan data ada
  if (poolPG.length > 0) selected.push(poolPG.pop());
  if (poolPGK.length > 0) selected.push(poolPGK.pop());
  if (poolBS.length > 0) selected.push(poolBS.pop());

  // Gabungkan sisa soal yang ada
  let remainingPool = [...poolPG, ...poolPGK, ...poolBS];
  shuffle(remainingPool);

  // Penuhi sisa kuota soal hingga mencapai targetTotal
  while (selected.length < targetTotal && remainingPool.length > 0) {
    selected.push(remainingPool.pop());
  }

  // Acak kembali urutan akhir soal agar jenis soal tersebar acak
  return shuffle(selected);
}

// 4. Menyimpan Hasil Jawaban ke Sheet (Jawaban-Murid)
function submitAnswers(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Jawaban-Murid');
  
  const { studentName, studentClass, userAnswers, questions, cheatStatus } = payload;
  
  let correctCount = 0;
  const totalQuestions = questions.length;

  questions.forEach((q, index) => {
    const uAns = userAnswers[index];
    if (!uAns) return;

    if (q.type === 'PG' || q.type === 'BS') {
      if (String(uAns).trim().toUpperCase() === String(q.answer).toUpperCase()) {
        correctCount++;
      }
    } else if (q.type === 'PGK') {
      const userArr = Array.isArray(uAns) ? uAns.sort().join(',') : String(uAns);
      const keyArr = String(q.answer).split(',').map(s => s.trim()).sort().join(',');
      if (userArr.toUpperCase() === keyArr.toUpperCase()) {
        correctCount++;
      }
    }
  });

  const finalScore = Math.round((correctCount / totalQuestions) * 100);
  const timestamp = new Date();

  sheet.appendRow([
    timestamp,
    studentName,
    studentClass,
    correctCount,
    totalQuestions,
    finalScore,
    cheatStatus
  ]);

  return {
    correctCount,
    totalQuestions,
    finalScore,
    cheatStatus
  };
}