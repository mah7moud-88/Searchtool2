Office.onReady(() => {

  const searchBtn = document.getElementById("searchBtn");
  const exportBtn = document.getElementById("exportBtn");
  const clearBtn = document.getElementById("clearBtn");
  const excelFile = document.getElementById("excelFile");
  const modeBtn = document.getElementById("modeBtn");

  // =========================
  // الأحداث
  // =========================

  if (excelFile) {
    excelFile.addEventListener(
      "change",
      importAccountsFromFile
    );
  }

  if (searchBtn) {
    searchBtn.onclick = searchAccount;
  }

  if (exportBtn) {
    exportBtn.onclick = exportExcel;
  }

  if (clearBtn) {

    clearBtn.onclick = () => {

      document.getElementById("accountNumber").value = "";
      document.getElementById("customerName").value = "";

      document.getElementById("result").innerText = "تم المسح";

      resultsData = [];
    };
  }


  // =========================
  // تغيير وضع البحث
  // =========================

  if (modeBtn) {

    modeBtn.onclick = () => {

      searchMode =
        searchMode === "independent"
          ? "paired"
          : "independent";

      modeBtn.innerText =
        searchMode === "independent"
          ? "🔁 وضع البحث: مستقل"
          : "🔗 وضع البحث: مطابق";
    };
  }

});


// ======================================================
// المتغيرات
// ======================================================

let searchMode = "independent";

let resultsData = [];

let accountIndex = {};
let nameIndex = {};
let pairIndex = {};

let last6Index = {};
let last5Index = {};


// ======================================================
// عدد الصفوف التي نقرأها في كل دفعة
// ======================================================

const CHUNK_SIZE = 5000;

const MAX_ROWS = 50000;


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

  accountIndex = {};
  nameIndex = {};
  pairIndex = {};

  last6Index = {};
  last5Index = {};
}


// ======================================================
// إضافة صف إلى الفهارس
// ======================================================

function addRowToIndex(
  row,
  rowIndex
) {

  // B = الاسم
  // C = الحساب
  // N = الملاحظة

  const name =
    row[0] ?? "";

  const account =
    row[1] ?? "";

  const cleanName =
    cleanValue(name);

  const cleanAcc =
    cleanValue(account);


  // ==================================================
  // الرقم الكامل
  // ==================================================

  if (cleanAcc) {

    if (
      accountIndex[cleanAcc] === undefined
    ) {

      accountIndex[cleanAcc] = [];
    }

    accountIndex[cleanAcc].push(rowIndex);
  }


  // ==================================================
  // الاسم
  // ==================================================

  if (cleanName) {

    if (!nameIndex[cleanName]) {

      nameIndex[cleanName] = [];
    }

    nameIndex[cleanName].push(rowIndex);
  }


  // ==================================================
  // الاسم + الرقم
  // ==================================================

  if (
    cleanName &&
    cleanAcc
  ) {

    const key =
      cleanName + "|" + cleanAcc;


    if (
      pairIndex[key] === undefined
    ) {

      pairIndex[key] = [];
    }

    pairIndex[key].push(rowIndex);
  }


  // ==================================================
  // آخر 6 أرقام
  // ==================================================

  if (
    cleanAcc &&
    cleanAcc.length >= 6
  ) {

    const last6 =
      cleanAcc.slice(-6);


    if (!last6Index[last6]) {

      last6Index[last6] = [];
    }


    last6Index[last6].push(rowIndex);
  }


  // ==================================================
  // آخر 5 أرقام
  // ==================================================

  if (
    cleanAcc &&
    cleanAcc.length >= 5
  ) {

    const last5 =
      cleanAcc.slice(-5);


    if (!last5Index[last5]) {

      last5Index[last5] = [];
    }


    last5Index[last5].push(rowIndex);
  }
}


// ======================================================
// البحث بالرقم
//
// كامل
// آخر 6
// آخر 5
// ======================================================

function findAccountRows(acc) {

  const cleanAcc =
    cleanValue(acc);


  if (!cleanAcc) {
    return [];
  }


  // ==================================================
  // الرقم الكامل
  // ==================================================

  if (
    accountIndex[cleanAcc] !== undefined
  ) {

    return accountIndex[cleanAcc];
  }


  // ==================================================
  // آخر 5 أرقام
  // ==================================================

  if (
    cleanAcc.length === 5
  ) {

    return (
      last5Index[cleanAcc] || []
    );
  }


  // ==================================================
  // آخر 6 أرقام
  // ==================================================

  if (
    cleanAcc.length === 6
  ) {

    return (
      last6Index[cleanAcc] || []
    );
  }


  // ==================================================
  // أكثر من 6 أرقام
  // يبحث بآخر 6
  // ==================================================

  if (
    cleanAcc.length > 6
  ) {

    const last6 =
      cleanAcc.slice(-6);


    return (
      last6Index[last6] || []
    );
  }


  // أقل من 5 أرقام
  return [];
}


// ======================================================
// رفع ملف Excel
// ======================================================

function importAccountsFromFile(e) {

  const file =
    e.target.files[0];


  if (!file) {
    return;
  }


  const reader =
    new FileReader();


  reader.onload = function (evt) {

    try {

      const data =
        new Uint8Array(
          evt.target.result
        );


      const workbook =
        XLSX.read(
          data,
          {
            type: "array"
          }
        );


      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ];


      const rows =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1
          }
        );


      let names = [];
      let accounts = [];


      rows.forEach(row => {

        if (row[0]) {

          names.push(
            String(row[0]).trim()
          );
        }


        if (row[1]) {

          accounts.push(
            String(row[1]).trim()
          );
        }

      });


      document.getElementById(
        "customerName"
      ).value =
        names.join("\n");


      document.getElementById(
        "accountNumber"
      ).value =
        accounts.join("\n");


      document.getElementById(
        "result"
      ).innerText =
        "✅ تم تحميل الملف";


    } catch (error) {

      console.error(error);

      document.getElementById(
        "result"
      ).innerText =
        "❌ حدث خطأ أثناء قراءة الملف";
    }

  };


  reader.readAsArrayBuffer(file);
}


// ======================================================
// البحث
// ======================================================

async function searchAccount() {

  const resultDiv =
    document.getElementById("result");


  const accountInput =
    document
      .getElementById("accountNumber")
      .value
      .trim();


  const nameInput =
    document
      .getElementById("customerName")
      .value
      .trim();


  const accounts =
    accountInput
      .split(/\r?\n/)
      .map(x => x.trim())
      .filter(Boolean);


  const names =
    nameInput
      .split(/\r?\n/)
      .map(x => x.trim())
      .filter(Boolean);


  if (
    accounts.length === 0 &&
    names.length === 0
  ) {

    resultDiv.innerText =
      "⚠️ اكتب رقم الحساب أو الاسم أولاً";

    return;
  }


  const totalCount =
    searchMode === "independent"
      ? accounts.length + names.length
      : Math.max(
          accounts.length,
          names.length
        );


  resultDiv.innerText =
    "🔄 جاري قراءة بيانات Excel...";


  resultsData = [];

  resetIndexes();


  try {

    await Excel.run(async (context) => {


      // ==================================================
      // قراءة Excel على دفعات
      // ==================================================

      for (
        let startRow = 0;
        startRow < MAX_ROWS;
        startRow += CHUNK_SIZE
      ) {

        const rowsToRead =
          Math.min(
            CHUNK_SIZE,
            MAX_ROWS - startRow
          );


        // ==================================================
        // B:C
        //
        // B = الاسم
        // C = الحساب
        // ==================================================

        const nameAccountRange =
          context
            .workbook
            .worksheets
            .getActiveWorksheet()
            .getRangeByIndexes(
              startRow,
              1,
              rowsToRead,
              2
            );


        // ==================================================
        // N
        //
        // N = الملاحظة
        // ==================================================

        const noteRange =
          context
            .workbook
            .worksheets
            .getActiveWorksheet()
            .getRangeByIndexes(
              startRow,
              13,
              rowsToRead,
              1
            );


        nameAccountRange.load("text");

        noteRange.load("text");


        await context.sync();


        const nameAccountValues =
          nameAccountRange.text;


        const noteValues =
          noteRange.text;


        // ==================================================
        // دمج B:C مع N
        // ==================================================

        for (
          let i = 0;
          i < rowsToRead;
          i++
        ) {

          const row = [

            nameAccountValues[i]?.[0] ?? "",

            nameAccountValues[i]?.[1] ?? "",

            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",

            noteValues[i]?.[0] ?? ""

          ];


          // الرقم الحقيقي للصف
          const actualRowIndex =
            startRow + i;


          addRowToIndex(
            row,
            actualRowIndex
          );
        }


        // ==================================================
        // تحديث الحالة
        // ==================================================

        const current =
          Math.min(
            startRow + rowsToRead,
            MAX_ROWS
          );


        resultDiv.innerText =
          `🔄 جاري قراءة البيانات... ${current.toLocaleString()} / ${MAX_ROWS.toLocaleString()}`;

      }


      // ==================================================
      // بدء البحث
      // ==================================================

      resultDiv.innerText =
        "🔍 جاري البحث...";


      let output = "";

      let foundCount = 0;


      // ==================================================
      // الوضع المستقل
      // ==================================================

      if (
        searchMode === "independent"
      ) {


        // ==================================================
        // البحث بالأرقام
        // ==================================================

        for (
          const acc of accounts
        ) {

          const rows =
            findAccountRows(acc);


          if (
            rows.length > 0
          ) {


            for (
              const rowIndex of rows
            ) {

              // نقرأ الصف المطلوب فقط
              // B:N

              const resultRange =
                context
                  .workbook
                  .worksheets
                  .getActiveWorksheet()
                  .getRangeByIndexes(
                    rowIndex,
                    1,
                    1,
                    13
                  );


              resultRange.load("text");

              await context.sync();


              const row =
                resultRange.text[0];


              const name =
                row[0] ?? "";


              const account =
                row[1] ?? "";


              const note =
                row[12] ?? "";


              const item = {

                name: name,

                account: account,

                note: note,

                status: "موجود"

              };


              resultsData.push(item);


              output +=
                `👤 ${name}\n` +
                `📌 ${account}\n` +
                `📝 ${note}\n\n`;


              foundCount++;
            }


          } else {


            resultsData.push({

              name: "",

              account: acc,

              note: "",

              status: "غير موجود"

            });


            output +=
              `❌ ${acc}\n\n`;
          }

        }


        // ==================================================
        // البحث بالأسماء
        // ==================================================

        for (
          const name of names
        ) {

          const cleanName =
            cleanValue(name);


          const rows =
            nameIndex[cleanName] || [];


          if (
            rows.length > 0
          ) {


            for (
              const rowIndex of rows
            ) {

              const resultRange =
                context
                  .workbook
                  .worksheets
                  .getActiveWorksheet()
                  .getRangeByIndexes(
                    rowIndex,
                    1,
                    1,
                    13
                  );


              resultRange.load("text");

              await context.sync();


              const row =
                resultRange.text[0];


              const resultName =
                row[0] ?? "";


              const account =
                row[1] ?? "";


              const note =
                row[12] ?? "";


              const item = {

                name: resultName,

                account: account,

                note: note,

                status: "موجود"

              };


              resultsData.push(item);


              output +=
                `👤 ${resultName}\n` +
                `📌 ${account}\n` +
                `📝 ${note}\n\n`;


              foundCount++;
            }


          } else {


            resultsData.push({

              name: name,

              account: "",

              note: "",

              status: "غير موجود"

            });


            output +=
              `❌ ${name}\n\n`;
          }

        }

      }


      // ==================================================
      // الوضع المطابق
      // ==================================================

      else {


        for (
          let i = 0;
          i < Math.max(
            accounts.length,
            names.length
          );
          i++
        ) {

          const acc =
            accounts[i] || "";


          const name =
            names[i] || "";


          if (
            !acc ||
            !name
          ) {

            resultsData.push({

              name: name,

              account: acc,

              note: "",

              status: "بيانات ناقصة"

            });


            output +=
              `⚠️ بيانات ناقصة\n` +
              `👤 ${name}\n` +
              `📌 ${acc}\n\n`;


            continue;
          }


          const rows =
            findAccountRows(acc);


          let matched =
            false;


          for (
            const rowIndex of rows
          ) {

            const resultRange =
              context
                .workbook
                .worksheets
                .getActiveWorksheet()
                .getRangeByIndexes(
                  rowIndex,
                  1,
                  1,
                  13
                );


            resultRange.load("text");

            await context.sync();


            const row =
              resultRange.text[0];


            const rowName =
              row[0] ?? "";


            const rowAccount =
              row[1] ?? "";


            const note =
              row[12] ?? "";


            if (
              cleanValue(rowName) ===
              cleanValue(name)
            ) {

              const item = {

                name: rowName,

                account: rowAccount,

                note: note,

                status: "موجود"

              };


              resultsData.push(item);


              output +=
                `👤 ${rowName}\n` +
                `📌 ${rowAccount}\n` +
                `📝 ${note}\n\n`;


              foundCount++;

              matched = true;

              break;
            }

          }


          if (!matched) {

            resultsData.push({

              name: name,

              account: acc,

              note: "",

              status: "غير مطابق"

            });


            output +=
              `❌ غير مطابق\n` +
              `👤 ${name}\n` +
              `📌 ${acc}\n\n`;
          }

        }

      }


      // ==================================================
      // عرض النتيجة
      // ==================================================

      resultDiv.innerText =
        `✅ تم العثور على ${foundCount} من أصل ${totalCount}\n\n` +
        output;

    });


  } catch (error) {

    console.error(error);


    resultDiv.innerText =
      "❌ حدث خطأ أثناء البحث: " +
      (
        error.message ||
        "خطأ غير معروف"
      );

  }

}


// ======================================================
// تصدير النتائج
// ======================================================

async function exportExcel() {

  const resultDiv =
    document.getElementById("result");


  if (
    !resultsData.length
  ) {

    resultDiv.innerText =
      "لا يوجد بيانات للتصدير";

    return;
  }


  try {

    await Excel.run(async (context) => {


      // ==================================================
      // إنشاء ورقة التصدير
      // ==================================================

      const sheet =
        context
          .workbook
          .worksheets
          .add("Export");


      // ==================================================
      // العناوين
      // ==================================================

      const data = [

        [
          "رقم الحساب",
          "الاسم",
          "الملاحظة",
          "الحالة"
        ]

      ];


      // ==================================================
      // إضافة النتائج
      // ==================================================

      resultsData.forEach(item => {

        data.push([

          item.account,

          item.name,

          item.note,

          item.status

        ]);

      });


      // ==================================================
      // كتابة البيانات
      // ==================================================

      const range =
        sheet.getRange(
          `A1:D${data.length}`
        );


      range.values =
        data;


      // ==================================================
      // تنسيق العناوين
      // ==================================================

      const header =
        sheet.getRange("A1:D1");


      header.format.font.bold =
        true;


      header.format.fill.color =
        "#D9EAF7";


      // ==================================================
      // ضبط الأعمدة
      // ==================================================

      range.format.autofitColumns();


      // ==================================================
      // تفعيل الورقة
      // ==================================================

      sheet.activate();


      await context.sync();

    });


    resultDiv.innerText =
      "✅ تم تصدير النتائج بنجاح";


  } catch (error) {

    console.error(error);


    resultDiv.innerText =
      "❌ حدث خطأ أثناء التصدير: " +
      (
        error.message ||
        "خطأ غير معروف"
      );
  }

}
