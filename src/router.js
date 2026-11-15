import { createRouter, createWebHashHistory } from 'vue-router';
import HomeView from './views/HomeView.vue';
import { currentDoc } from './store.js';

// Hash history works on GitHub Pages without any server-side rewrites.
const router = createRouter({
  history: createWebHashHistory(),
  routes: [
    { path: '/', name: 'home', component: HomeView },
    {
      path: '/edit',
      name: 'edit',
      component: () => import('./views/EditorView.vue'),
      beforeEnter: () => (currentDoc.value ? true : { name: 'home' }),
    },
    { path: '/about', name: 'about', component: () => import('./views/AboutView.vue') },
    { path: '/:pathMatch(.*)*', redirect: '/' },
  ],
});

router.afterEach((to) => {
  document.title = to.name === 'edit' && currentDoc.value ? `${currentDoc.value.name} – Intact` : 'Intact — private .docx editor';
});

export default router;
