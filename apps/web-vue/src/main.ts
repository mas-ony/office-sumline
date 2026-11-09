import { createApp } from 'vue';
import '@sumline/vue/style.css';
import '@sumline/web-shared/app.css';
import App from './App.vue';
import { registerServiceWorker } from '@sumline/web-shared/shell';

createApp(App).mount('#root');

if (import.meta.env.PROD) registerServiceWorker();
