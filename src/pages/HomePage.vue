<template>
  <div class="home">
    <section class="heading">
      <h1>Export & Import Functions</h1>
    </section>

  </div>
</template>

<script setup>
import { ref } from "vue";

const userData = ref([
  {
    user_id: "1234_MH",
    email: "minhua@acxhk.com",
    password: "Qazwsx123",
    user_type: "Factory",
    create_date: "2025/01/07 22:20:48",
  },
  {
    user_id: "3456_Yo",
    email: "yo@test.com",
    password: "Qazwsx123",
    user_type: "Buyer",
    create_date: "2025/01/07 22:20:48",
  },
  {
    user_id: "3456_Yo",
    email: "yo@test.com",
    password: "Qazwsx123",
    user_type: "Buyer",
    create_date: "2025/01/07 22:20:48",
  },
  {
    user_id: "3456_Yo",
    email: "yo@test.com",
    password: "Qazwsx123",
    user_type: "Buyer",
    create_date: "2025/01/07 22:20:48",
  },
]);

let importedUserData = ref([]);
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

.left, .right {
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

.export-btn, .import-btn {
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
