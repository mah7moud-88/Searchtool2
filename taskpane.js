Office.onReady(() => {

  const searchBtn      = document.getElementById("searchBtn");
  const exportBtn      = document.getElementById("exportBtn");
  const clearBtn       = document.getElementById("clearBtn");
  const excelFile      = document.getElementById("excelFile");
  const modeBtn        = document.getElementById("modeBtn");
  const fieldBtn       = document.getElementById("fieldBtn");
  const fieldMenu      = document.getElementById("fieldMenu");
  const clearFilterBtn = document.getElementById("clearFilterBtn");

  // =========================
  // الأحداث
  // =========================

  if (excelFile) {
    excelFile.addEventListener("change", importAccountsFromFile);
  }

  if (searchBtn) {
    searchBtn.onclick = () => {
      if (searchButtonMode === "search") {
        searchAccount();
      } else {
        applyFilterToSheet();
      }
    };
  }

  if (exportBtn) exportBtn.onclick = exportExcel;

  if (clearBtn) {
    clearBtn.onclick = () => {
      document.getElementById("accountNumber").value = "";
      document.getElementById("customerName").value = "";
      document.getElementById("result").innerText = "تم المسح";
      resultsData = [];
      resetIndexes();

      searchButtonMode = "search";
      searchBtn.innerText = "🔎 بحث";
      lastFilterValues = [];
    };
  }

  if (clearFilterBtn) {
    clearFilterBtn.onclick = async () => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.autoFilter.remove();
          await context.sync();
        });
        document.getElementById("result").innerText = "✅ تم إلغاء الفلتر";
      } catch (error) {
        console.error(error);
        document.getElementById("result").innerText =
          "❌ خطأ أثناء إلغاء الفلتر: " + (error.message || "");
      }
    };
  }

  if (modeBtn) {
    modeBtn.onclick = () => {
      searchMode = searchMode === "independent" ? "paired" : "independent";
      modeBtn.innerText =
        searchMode === "independent"
          ? "🔁 وضع البحث: مستقل"
          : "🔗 وضع البحث: مطابق";
    };
  }

  // =========================
  // زر "حدد طريقة البحث"
  // =========================

  const FIELD_LABELS = {
    account:  "🔎 رقم الحساب",
    phone:    "🔎 رقم الهاتف",
    national: "🔎 الرقم الوطني",
    passport: "🔎 رقم الجواز"
  };

  const FIELD_TITLES = {
    account:  "رقم الحساب",
    phone:    "رقم الهاتف",
    national: "الرقم الوطني",
    passport: "رقم الجواز"
  };

  const PLACEHOLDERS = {
    account:  "اكتب أرقام الحسابات كل رقم في سطر",
    phone:    "اكتب أرقام الهواتف كل رقم في سطر",
    national: "اكتب الأرقام الوطنية كل رقم في سطر",
    passport: "اكتب أرقام الجوازات كل رقم في سطر"
  };

  if (fieldBtn && fieldMenu) {

    fieldBtn.onclick = (e) => {
      e.stopPropagation();
      fieldMenu.style.display =
        fieldMenu.style.display === "block" ? "none" : "block";
    };

    fieldMenu.querySelectorAll(".field-item").forEach(item => {
      item.onclick = () => {
        searchField = item.dataset.field;

        fieldBtn.innerText = FIELD_LABELS[searchField];

        const accInput = document.getElementById("accountNumber");
        const accLabel = document.getElementById("accountLabel");

        if (accInput) accInput.placeholder = PLACEHOLDERS[searchField];
        if (accLabel) accLabel.innerText   = FIELD_TITLES[searchField];

        fieldMenu.style.display = "none";
      };
    });

    document.addEventListener("click", () => {
      fieldMenu.style.display = "none";
    });
  }

});


// ======================================================
// المتغيرات
// ======================================================

let searchMode  = "independent";
let searchField = "account";
let resultsData = [];

let accountIndex  = {};
let nameIndex     = {};
let pairIndex     = {};

let last6Index = {};
let last5Index = {};

let nationalIndex = {};
let phoneIndex    = {};
let passportIndex = {};

// للفلترة
let searchButtonMode = "search";
let lastFilterValues = [];

const CHUNK_SIZE = 5000;
const MAX_ROWS   = 50000;


// ======================================================
// تنظيف البيانات
// ======================================================

function cleanValue(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}


// ======================================================
// تصفير الفهارس
// ======================================================

function resetIndexes() {
  accountIndex  = {};
  nameIndex     = {};
  pairIndex     = {};
  last6Index    = {};
  last5Index    = {};
  nationalIndex = {};
  phoneIndex    = {};
  passportIndex = {};
}


// ======================================================
// إضافة صف إلى الفهارس
// B=الاسم(0) C=الحساب(1) D=الجواز(2) E=فاضي(3)
// F=الوطني(4) G=الهاتف(5)
// ======================================================

function addRowToIndex(row, rowIndex) {

  const name     = row[0] ?? "";
  const account  = row[1] ?? "";
  const passport = row[2] ?? "";
  const national = row[4] ?? "";
  const phone    = row[5] ?? "";

  const cleanName     = cleanValue(name);
  const cleanAcc      = cleanValue(account);
  const cleanPassport = cleanValue(passport);
  const cleanNational = cleanValue(national);
  const cleanPhone    = cleanValue(phone);

  if (cleanName) {
    if (!nameIndex[cleanName]) nameIndex[cleanName] = [];
    nameIndex[cleanName].push(rowIndex);
  }

  if (cleanAcc) {
    if (accountIndex[cleanAcc] === undefined) accountIndex[cleanAcc] = [];
    accountIndex[cleanAcc].push(rowIndex);
  }

  if (cleanName && cleanAcc) {
    const key = cleanName + "|" + cleanAcc;
    if (pairIndex[key] === undefined) pairIndex[key] = [];
    pairIndex[key].push(rowIndex);
  }

  if (cleanAcc && cleanAcc.length >= 6) {
    const last6 = cleanAcc.slice(-6);
    if (!last6Index[last6]) last6Index[last6] = [];
    last6Index[last6].push(rowIndex);
  }

  if (cleanAcc && cleanAcc.length >= 5) {
    const last5 = cleanAcc.slice(-5);
    if (!last5Index[last5]) last5Index[last5] = [];
    last5Index[last5].push(rowIndex);
  }

  if (cleanPassport) {
    if (!passportIndex[cleanPassport]) passportIndex[cleanPassport] = [];
    passportIndex[cleanPassport].push(rowIndex);
  }

  if (cleanNational) {
    if (!nationalIndex[cleanNational]) nationalIndex[cleanNational] = [];
    nationalIndex[cleanNational].push(rowIndex);
  }

  if (cleanPhone) {
    if (!phoneIndex[cleanPhone]) phoneIndex[cleanPhone] = [];
    phoneIndex[cleanPhone].push(rowIndex);
  }
}


// ======================================================
// البحث حسب الحقل المختار
// ======================================================

function findRows(value) {

  const v = cleanValue(value);
  if (!v) return [];

  switch (searchField) {

    case "phone":
      return phoneIndex[v] || [];

    case "national":
      return nationalIndex[v] || [];

    case "passport":
      return passportIndex[v] || [];

    case "account":
    default:
      if (accountIndex[v] !== undefined) return accountIndex[v];
      if (v.length === 5) return last5Index[v] || [];
      if (v.length === 6) return last6Index[v] || [];
      if (v.length > 6)   return last6Index[v.slice(-6)] || [];
      return [];
  }
}


// ======================================================
// جلب قيمة الحقل المختار من صف
// ======================================================

function getFieldValue(item) {
  switch (searchField) {
    case "phone":    return item.phone;
    case "national": return item.national;
    case "passport": return item.card;
    case "account":
    default:         return item.account;
  }
}


// ======================================================
// رفع ملف Excel
// ======================================================

function importAccountsFromFile(e) {

  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      let names = [];
      let accounts = [];

      rows.forEach(row => {
        if (row[0]) names.push(String(row[0]).trim());
        if (row[1]) accounts.push(String(row[1]).trim());
      });

      document.getElementById("customerName").value = names.join("\n");
      document.getElementById("accountNumber").value = accounts.join("\n");

      document.getElementById("result").innerText = "✅ تم تحميل الملف";

    } catch (error) {
      console.error(error);
      document.getElementById("result").innerText =
        "❌ حدث خطأ أثناء قراءة الملف";
    }
  };

  reader.readAsArrayBuffer(file);
}


// ======================================================
// البحث
// ======================================================

async function searchAccount() {

  const resultDiv = document.getElementById("result");

  const accountInput = document
    .getElementById("accountNumber").value.trim();

  const nameInput = document
    .getElementById("customerName").value.trim();

  const accounts = accountInput
    .split(/\r?\n/).map(x => x.trim()).filter(Boolean);

  const names = nameInput
    .split(/\r?\n/).map(x => x.trim()).filter(Boolean);

  if (accounts.length === 0 && names.length === 0) {
    resultDiv.innerText = "⚠️ اكتب رقم الحساب أو الاسم أولاً";
    return;
  }

  const totalCount =
    searchMode === "independent"
      ? accounts.length + names.length
      : Math.max(accounts.length, names.length);

  resultDiv.innerText = "🔄 جاري قراءة بيانات Excel...";

  resultsData = [];
  resetIndexes();

  searchButtonMode = "search";
  document.getElementById("searchBtn").innerText = "🔎 بحث";

  try {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // ==================================================
      // قراءة Excel على دفعات (B إلى G)
      // ==================================================

      for (let startRow = 0; startRow < MAX_ROWS; startRow += CHUNK_SIZE) {

        const rowsToRead = Math.min(CHUNK_SIZE, MAX_ROWS - startRow);

        const range = sheet.getRangeByIndexes(startRow, 1, rowsToRead, 6);
        range.load("text");

        await context.sync();

        const values = range.text;

        for (let i = 0; i < rowsToRead; i++) {

          const row = [
            values[i]?.[0] ?? "",
            values[i]?.[1] ?? "",
            values[i]?.[2] ?? "",
            values[i]?.[3] ?? "",
            values[i]?.[4] ?? "",
            values[i]?.[5] ?? ""
          ];

          addRowToIndex(row, startRow + i);
        }

        const current = Math.min(startRow + rowsToRead, MAX_ROWS);
        resultDiv.innerText =
          `🔄 جاري قراءة البيانات... ${current.toLocaleString()} / ${MAX_ROWS.toLocaleString()}`;
      }


      // ==================================================
      // بدء البحث
      // ==================================================

      resultDiv.innerText = "🔍 جاري البحث...";

      let output = "";
      let foundCount = 0;

      async function readFullRow(rowIndex) {
        const r = sheet.getRangeByIndexes(rowIndex, 1, 1, 6);
        r.load("text");
        await context.sync();

        const row = r.text[0];
        return {
          rowIndex,
          name:     row[0] ?? "",
          account:  row[1] ?? "",
          card:     row[2] ?? "",
          national: row[4] ?? "",
          phone:    row[5] ?? ""
        };
      }

      function formatOutput(item) {
        return (
          `👤 ${item.name}\n` +
          `📌 ${item.account}\n` +
          `🪪 ${item.card}\n` +
          `🆔 ${item.national}\n` +
          `📞 ${item.phone}\n\n`
        );
      }


      // ==================================================
      // الوضع المستقل
      // ==================================================

      if (searchMode === "independent") {

        for (const acc of accounts) {

          const rows = findRows(acc);

          if (rows.length > 0) {
            for (const rowIndex of rows) {

              const data = await readFullRow(rowIndex);

              const item = { ...data, status: "موجود" };
              resultsData.push(item);
              output += formatOutput(item);
              foundCount++;
            }
          } else {
            resultsData.push({
              rowIndex: -1,
              name: "", account: acc, card: "", national: "", phone: "",
              status: "غير موجود"
            });
            output += `❌ ${acc}\n\n`;
          }
        }

        for (const name of names) {

          const cleanName = cleanValue(name);
          const rows = nameIndex[cleanName] || [];

          if (rows.length > 0) {
            for (const rowIndex of rows) {

              const data = await readFullRow(rowIndex);

              const item = { ...data, status: "موجود" };
              resultsData.push(item);
              output += formatOutput(item);
              foundCount++;
            }
          } else {
            resultsData.push({
              rowIndex: -1,
              name: name, account: "", card: "", national: "", phone: "",
              status: "غير موجود"
            });
            output += `❌ ${name}\n\n`;
          }
        }
      }


      // ==================================================
      // الوضع المطابق
      // ==================================================

      else {

        for (let i = 0; i < Math.max(accounts.length, names.length); i++) {

          const acc  = accounts[i] || "";
          const name = names[i] || "";

          if (!acc || !name) {
            resultsData.push({
              rowIndex: -1,
              name, account: acc, card: "", national: "", phone: "",
              status: "بيانات ناقصة"
            });
            output += `⚠️ بيانات ناقصة\n👤 ${name}\n📌 ${acc}\n\n`;
            continue;
          }

          const rows = findRows(acc);
          let matched = false;

          for (const rowIndex of rows) {

            const data = await readFullRow(rowIndex);

            if (
              cleanValue(data.name) === cleanValue(name) &&
              cleanValue(getFieldValue(data)) === cleanValue(acc)
            ) {

              const item = { ...data, status: "موجود" };
              resultsData.push(item);
              output += formatOutput(item);
              foundCount++;
              matched = true;
              break;
            }
          }

          if (!matched) {
            resultsData.push({
              rowIndex: -1,
              name, account: acc, card: "", national: "", phone: "",
              status: "غير مطابق"
            });
            output += `❌ غير مطابق\n👤 ${name}\n📌 ${acc}\n\n`;
          }
        }
      }


      // ==================================================
      // عرض النتيجة
      // ==================================================

      resultDiv.innerText =
        `✅ تم العثور على ${foundCount} من أصل ${totalCount}\n\n` + output;


      // ==================================================
      // تجهيز بيانات الفلتر
      // ==================================================

      const rawValues = accountInput
        .split(/\r?\n/)
        .map(x => String(x).trim())
        .filter(v => v.length > 0);

      lastFilterValues = rawValues;

      const hasResults = resultsData.some(r => r.status === "موجود");

      if (hasResults && rawValues.length > 0) {
        searchButtonMode = "filter";
        document.getElementById("searchBtn").innerText = "🎯 توجيه";
      } else {
        searchButtonMode = "search";
        document.getElementById("searchBtn").innerText = "🔎 بحث";
      }

    });

  } catch (error) {

    console.error(error);
    resultDiv.innerText =
      "❌ حدث خطأ أثناء البحث: " + (error.message || "خطأ غير معروف");

    searchButtonMode = "search";
    document.getElementById("searchBtn").innerText = "🔎 بحث";
  }
}


// ======================================================
// تطبيق AutoFilter على الشيت - الإصدار النهائي
// ======================================================

async function applyFilterToSheet() {

  const resultDiv = document.getElementById("result");

  if (!lastFilterValues || !lastFilterValues.length) {
    resultDiv.innerText = "⚠️ لا توجد قيم للفلترة";
    return;
  }

  // ✅ التحويل الإجباري إلى String + تنظيف شامل
  const filterValues = lastFilterValues
    .map(v => String(v ?? "").trim())
    .filter(v => v.length > 0);

  if (!filterValues.length) {
    resultDiv.innerText = "⚠️ لا توجد قيم صالحة للفلترة";
    return;
  }

  try {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const used = sheet.getUsedRange();
      used.load(["rowIndex", "rowCount"]);
      await context.sync();

      // مسح أي فلتر قديم
      try {
        sheet.autoFilter.remove();
        await context.sync();
      } catch (_) {}

      // نطاق الفلتر (B إلى G) - يشمل الهيدر
      const filterRange = sheet.getRangeByIndexes(
        used.rowIndex,
        1,
        used.rowCount,
        6
      );

      // فهرس العمود النسبي داخل النطاق (B=0, C=1, D=2, F=4, G=5)
      const FIELD_TO_REL_COL = {
        account:  1,  // C
        passport: 2,  // D
        national: 4,  // F
        phone:    5   // G
      };
      const colInRange = FIELD_TO_REL_COL[searchField] ?? 1;

      // ✅ التطبيق الصحيح مع قيم نصية نظيفة
      sheet.autoFilter.apply(filterRange, colInRange, {
        filterOn: Excel.FilterOn.values,
        values: filterValues
      });

      sheet.activate();
      await context.sync();

      resultDiv.innerText =
        `🎯 تم توجيه الشيت إلى ${filterValues.length} قيمة.\n` +
        `لإلغاء الفلتر: اضغط "🚫 إلغاء الفلتر" أو من أيقونة الفلتر في الشيت.`;

      searchButtonMode = "search";
      document.getElementById("searchBtn").innerText = "🔎 بحث";
      lastFilterValues = [];

    });

  } catch (error) {
    console.error(error);
    resultDiv.innerText =
      "❌ حدث خطأ أثناء التوجيه: " + (error.message || "خطأ غير معروف");

    searchButtonMode = "search";
    document.getElementById("searchBtn").innerText = "🔎 بحث";
  }
}


// ======================================================
// تصدير النتائج
// ======================================================

async function exportExcel() {

  const resultDiv = document.getElementById("result");

  if (!resultsData.length) {
    resultDiv.innerText = "لا يوجد بيانات للتصدير";
    return;
  }

  try {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.add("Export");

      const data = [[
        "رقم الحساب",
        "الاسم",
        "رقم الجواز",
        "الرقم الوطني",
        "رقم الهاتف",
        "الحالة"
      ]];

      resultsData.forEach(item => {
        data.push([
          item.account,
          item.name,
          item.card,
          item.national,
          item.phone,
          item.status
        ]);
      });

      const range = sheet.getRange(`A1:F${data.length}`);
      range.values = data;

      const header = sheet.getRange("A1:F1");
      header.format.font.bold = true;
      header.format.fill.color = "#D9EAF7";

      range.format.autofitColumns();
      sheet.activate();

      await context.sync();
    });

    resultDiv.innerText = "✅ تم تصدير النتائج بنجاح";

  } catch (error) {
    console.error(error);
    resultDiv.innerText =
      "❌ حدث خطأ أثناء التصدير: " + (error.message || "خطأ غير معروف");
  }
}
