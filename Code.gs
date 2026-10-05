const SS = SpreadsheetApp.getActiveSpreadsheet();
const CASE_SHEET = SS.getSheetByName("KasusPembelajaran");
const RESPONSE_SHEET = SS.getSheetByName("Response") || SS.insertSheet("Response");

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("E-LKPD Kinematika GLB dan GLBB")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initResponseSheet() {
  if (!RESPONSE_SHEET) return;
  const firstRow = RESPONSE_SHEET.getRange(1, 1, 1, 6).getValues()[0];
  if (firstRow[0] && firstRow[0] === "Timestamp") {
    return;
  }

  RESPONSE_SHEET.clear();
  RESPONSE_SHEET.getRange(1, 1, 1, 6).setValues([
    ["Timestamp", "Kelas", "Kelompok", "Nama Siswa", "Nama Anggota", "Jawaban JSON"]
  ]);
  RESPONSE_SHEET.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1F4E78").setFontColor("white");
  RESPONSE_SHEET.setFrozenRows(1);
  RESPONSE_SHEET.setColumnWidths(1, 6, [150, 80, 80, 180, 220, 500]);
}

function getCaseByGroup(group) {
  if (!CASE_SHEET) return null;
  const values = CASE_SHEET.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(group).trim()) {
      const row = values[i];
      return {
        sekolah: "SMAN 1 Asembagus",
        judul: "Kinematika GLB dan GLBB",
        tahunAjaran: "2026/2027",
        narasi: row[1] || "",
        sections: [
          { title: "MASALAH", questions: [row[2], row[3], row[4]] },
          { title: "PENYEBAB", questions: [row[5], row[6], row[7]] },
          { title: "DAMPAK", questions: [row[8], row[9], row[10]] },
          { title: "PIHAK TERLIBAT", questions: [row[11], row[12], row[13]] },
          { title: "SOLUSI", questions: [row[14], row[15], row[16]] },
          { title: "INDIKATOR KEBERHASILAN", questions: [row[17], row[18], row[19]] }
        ]
      };
    }
  }

  return null;
}

function saveResponse(payload) {
  if (!payload || !payload.kelas || !payload.kelompok || !payload.namaSiswa) {
    return "Data siswa belum lengkap. Isi Kelas, Kelompok, dan Nama Siswa terlebih dahulu.";
  }

  initResponseSheet();

  const row = [
    new Date(),
    payload.kelas,
    payload.kelompok,
    payload.namaSiswa,
    payload.namaAnggota || "",
    JSON.stringify(payload.answers || {})
  ];

  RESPONSE_SHEET.appendRow(row);
  return "✅ Jawaban berhasil disimpan ke Google Sheets!";
}
