<template>
  <div id="app">
    <HelloWorld msg="Welcome to Your Vue.js App"/>
  </div>
</template>

<script>
import HelloWorld from './components/HelloWorld.vue'

export default {
  name: 'App',
  components: {
    HelloWorld
  }
}
</script>

<style>
html {
  overflow-y: scroll;
  font-family: helvetica, arial, sans-serif;
}

body {
  margin: 0;
  padding: 0;
}
</style>
