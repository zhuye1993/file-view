import Vue from 'vue'
import App from './App.vue'
import VueViewer from 'v-viewer';
import 'viewerjs/dist/viewer.css'

Vue.config.productionTip = false

Vue.use(VueViewer)

new Vue({
  render: h => h(App),
}).$mount('#app')
