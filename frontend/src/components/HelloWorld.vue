<template>
  <div class="hello">
    <h1>{{ msg }}</h1>
    
    <h2> Exchange latest rates </h2>
    <center>
      <table id="tableLastData">
        <thead>
          <tr>
            <th>Date</th>
            <th>EUR</th>
            <th>USD</th>
            <th>JPY</th>
            <th>GBP</th>
          </tr>
        </thead>
        <tbody>
            <td>{{ json_data.date }}</td>
            <td>{{ json_data.eur }}</td>
            <td>{{ json_data.usd }}</td>
            <td>{{ json_data.jpy }}</td>
            <td>{{ json_data.gbp }}</td>
        </tbody>
      </table> <br>
      <button v-on:click="loadLastData">Update</button>
      </center>

      <h3> Exchange all rates </h3>
      <center>
        <table id="tableAllData">
          <thead>
            <tr>
              <th>Date</th>
              <th>EUR</th>
              <th>USD</th>
              <th>JPY</th>
              <th>GBP</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="jad in json_all_data">
              <td>{{ jad.date }}</td>
              <td>{{ jad.eur }}</td>
              <td>{{ jad.usd }}</td>
              <td>{{ jad.jpy }}</td>
              <td>{{ jad.gbp }}</td>
            </tr> 
          </tbody>
        </table> <br>
        <button v-on:click="loadAllData">Update</button>
        </center>

      <h4> Write data to file </h4>
        <button v-on:click="saveDataToFile">Save data to file</button>

      <h4> Write data from file </h4>
        <button v-on:click="loadDataFromFile">Load data from file</button>

      <h4> Clear data in database </h4>
        <button v-on:click="clearDataInDatabase">Clear</button>

  </div>
</template>

<script>
import axios from 'axios'
export default {
  name: 'HelloWorld',
  data () {
    return {
      msg: 'Welcome to my test Vue.js App',
      json_data: [],
      json_all_data: []
    }
  },
  methods: {
    loadLastData() {
      axios
        .get('http://127.0.0.1:5000/currency', {headers: {'Content-Type': 'application/json'}})
        .then(response => (this.json_data = response.data))
    },
    loadAllData() {
      axios
        .get('http://127.0.0.1:5000/currencyAll', {headers: {'Content-Type': 'application/json'}})
        .then(response => (this.json_all_data = response.data))
    },
    saveDataToFile() {
      axios
        .get('http://127.0.0.1:5000/saveToFile')
    },
    loadDataFromFile() {
      axios
        .get('http://127.0.0.1:5000/readFromFile')
    },
    clearDataInDatabase() {
      axios
        .get('http://127.0.0.1:5000/clearDatabase')
    }
  },
  beforeMount(){
    this.loadLastData();
    this.loadAllData();
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h1, h2 {
  font-weight: normal;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
table {
  border-style: solid;
}
</style>
