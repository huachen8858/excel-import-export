<template>
  <div class="home">
    <section class="heading">
      <h1>Export & Import Functions</h1>
    </section>
    <section class="container">
      <!-- Export Function -->
      <div class="left">
        <div class="top">
          <h3>Export Data:</h3>
          <div class="data-wrapper">
            <pre><code>{{ userData }}</code></pre>
          </div>
        </div>
        <div class="bottom">
          <button class="export-btn" @click="exportExcel()">
            Export Excel File
          </button>
          <button class="export-btn" @click="exportCSV()">
            Export CSV File
          </button>
        </div>
      </div>

      <div class="divider"></div>

      <!-- Import Function -->
      <div class="right">
        <div class="top">
          <h3>Import Data:</h3>
          <div class="data-wrapper">
            <pre><code>{{ importedUserData }}</code></pre>
          </div>
        </div>
        <div class="bottom">
          <button class="import-btn" @click="fileInput.click">
            Import Excel File
          </button>
          <button class="import-btn" @click="csvFileInput.click">
            Import CSV File
          </button>
        </div>
        <input
          style="visibility: hidden"
          type="file"
          ref="fileInput"
          accept=".xlsx, .xls"
          @change="importExcel"
        />
        <input
          style="visibility: hidden"
          type="file"
          ref="csvFileInput"
          accept=".csv"
          @change="importCSV"
        />
      </div>
    </section>
  </div>
</template>

<script setup>
import { ref } from "vue";
import dayjs from "dayjs";
import ExcelJS from "exceljs";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const fileInput = ref(null);
const csvFileInput = ref(null);

const userExcelHeading = ref([
  { title: "User ID", key: "user_id" },
  { title: "User Name", key: "user_name" },
  { title: "Email", key: "email" },
  { title: "User Type", key: "user_type" },
  { title: "Gender", key: "gender" },
  { title: "Created Date", key: "create_date" },
]);

const userData = ref([
  {
    user_id: "1234",
    user_name: "Marry",
    email: "marry@test.com",
    user_type: "Admin",
    gender: "Female",
    create_date: "2025/01/07 22:20:48",
  },
  {
    user_id: "6666",
    user_name: "Yen",
    email: "yo@test.com",
    user_type: "User",
    gender: "Male",
    create_date: "2025/01/07 22:20:48",
  },
  {
    user_id: "3333",
    user_name: "Peter",
    email: "peter@test.com",
    user_type: "User",
    gender: "Male",
    create_date: "2025/01/07 22:20:48",
  },
  {
    user_id: "3456",
    user_name: "Zed",
    email: "zed@test.com",
    user_type: "User",
    gender: "Male",
    create_date: "2025/01/07 22:20:48",
  },
]);

let importedUserData = ref([]);

const exportExcel = async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("User");

  const columns = userExcelHeading.value.map((header) => ({
    header: header.title,
    key: header.key,
    style: { alignment: { horizontal: "center" } },
  }));
  worksheet.columns = columns;

  userData.value.forEach((row) => {
    worksheet.addRow(row);
  });

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const downloadLink = document.createElement("a");
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = "User.xlsx";

  downloadLink.click();
};

const exportCSV = () => {
  var tableData = [];
  var obj = {};

  userData.value.forEach((data) => {
    userExcelHeading.value.forEach((header) => {
      obj[header.title] = `${data[header.key]}`;
    });
    tableData.push(obj);
    obj = {};
  });

  const csvData = Papa.unparse({
    fields: userExcelHeading.value.title,
    data: tableData,
  });
  const blob = new Blob([csvData], { type: "text/csv;charset=utf-8;" });
  const anchor = document.createElement("a");
  const url = URL.createObjectURL(blob);

  anchor.setAttribute("href", url);
  anchor.setAttribute("download", "User.csv");

  anchor.click();

  URL.revokeObjectURL(url);
  anchor.remove();
};

const importExcel = (event) => {
  const file = event.target.files[0];

  if (!file || !file.name.match(/\.(xlsx|xls)$/)) {
    return; // If the file format is wrong, return directly
  }

  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Read the first Sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse into JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const processedData = [];

    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];

      if (row[0] && row[1]) {
        const user = {
          user_id: row[0],
          user_name: row[1],
          email: row[2],
          user_type: row[3] ? row[3] : "User",
          gender: row[4] ? row[4] : "Male",
          create_date: dayjs().format("YYYY/MM/DD HH:mm:ss"),
        };

        processedData.push(user);
      }
    }

    importedUserData.value = processedData;
  };

  reader.readAsArrayBuffer(file); // Read the file as binary data
};

const importCSV = (event) => {
  const file = event.target.files[0];

  if (!file || !file.name.match(/\.(csv)$/)) {
    return; // If the file format is wrong, return directly
  }

  const reader = new FileReader();

  reader.onload = (e) => {
    const csvData = Papa.parse(e.target.result, {
      header: true,
      skipEmptyLines: true,
    });

    const processedData = csvData.data.map((row) => ({
      user_id: row["User ID"],
      user_name: row["User Name"],
      email: row["Email"],
      user_type: row["User Type"] ? row["User Type"] : "User",
      gender: row["Gender"] ? row["Gender"] : "Male",
      create_date: dayjs().format("YYYY/MM/DD HH:mm:ss"),
    }));

    importedUserData.value = processedData;
  };

  reader.readAsText(file); // Read the file as text
};
</script>

<style scoped>
.home {
  width: 100%;
  height: 96%;
}

.heading {
  width: 100%;
  height: 10%;
  display: flex;
  justify-content: center;
  align-items: center;
}

.container {
  width: 90%;
  height: 86%;
  padding: 2% 5%;
  display: flex;
  justify-content: space-evenly;
}

.divider {
  width: 1px;
  height: 100%;
  background-color: #ccc;
}

.left,
.right {
  width: 40%;
}

.top {
  width: 100%;
  height: 90%;
}

.bottom {
  width: 100%;
  height: 10%;
  display: flex;
  align-items: center;
  justify-content: space-around;
}

.export-btn,
.import-btn {
  width: 10vw;
  height: 6vh;
  background-color: #45bbab;
  color: #fff;
  letter-spacing: 0.05vw;
  padding: 1% 2%;
  border: none;
  border-radius: 5px;
  cursor: pointer;
}

.import-btn {
  background-color: #ffbb00;
}

.back-btn {
  width: 14vw;
  height: 6vh;
  background-color: #9b9b9b;
  color: #fff;
  letter-spacing: 0.05vw;
  padding: 1% 2%;
  border: none;
  border-radius: 5px;
  cursor: pointer;
}

.export-btn:hover,
.import-btn:hover {
  box-shadow: 4px 4px 10px rgba(120, 120, 120, 0.1);
}

.data-wrapper {
  width: 100%;
  height: 90%;
  overflow: scroll;
}
</style>
