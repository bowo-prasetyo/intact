<template>
  <header class="app-header">
    <RouterLink to="/" class="brand">
      <img :src="icon" alt="" width="24" height="24" />
      <span>Intact</span>
    </RouterLink>
    <slot />
    <nav class="nav">
      <RouterLink to="/about">About</RouterLink>
    </nav>
  </header>
</template>

<script setup>
const icon = `${import.meta.env.BASE_URL}favicon.svg`;
</script>

<style scoped>
.app-header {
  display: flex; align-items: center; gap: 16px; padding: 8px 16px; min-height: 48px;
  background: var(--brand); color: #fff;
}
.brand { display: flex; align-items: center; gap: 8px; color: inherit; text-decoration: none; font-weight: 600; font-size: 1.05rem; }
.brand img { background: #fff; border-radius: 4px; }
.nav { margin-left: auto; }
.nav a { color: #fff; opacity: .9; text-decoration: none; }
.nav a:hover { opacity: 1; text-decoration: underline; }
</style>
