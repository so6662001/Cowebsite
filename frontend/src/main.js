import { createApp } from 'vue'
import { createUnhead } from '@unhead/vue'
import App from './App.vue'
import router from './router'
import './style.css'

const app = createApp(App)
createUnhead()

app.use(router)
app.mount('#app')
