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
          <button class="import-btn" @click="importExcel()">Import File</button>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
import { ref } from "vue";
import ExcelJS from "exceljs";

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
    user_id: "1234",
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
