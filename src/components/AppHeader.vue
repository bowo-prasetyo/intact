<template>
  <header class="app-header">
    <RouterLink to="/" class="brand" aria-label="Intact, start page">
      <!-- The Intact mark (public/logo.svg), reversed for the teal bar: a page with a text cursor. -->
      <svg class="mark" viewBox="0 0 64 64" width="26" height="26" aria-hidden="true">
        <path class="page" d="M18 4H38L54 20V52A8 8 0 0 1 46 60H18A8 8 0 0 1 10 52V12A8 8 0 0 1 18 4Z" />
        <path class="fold" d="M38 4V14A6 6 0 0 0 44 20H54Z" />
        <path class="beam" d="M24 18Q32 18 32 24Q32 18 40 18M32 24V40M24 46Q32 46 32 40Q32 46 40 46" />
      </svg>
      <span>Intact</span>
    </RouterLink>
    <slot />
    <nav class="nav">
      <RouterLink to="/about">About</RouterLink>
    </nav>
  </header>
</template>

<style scoped>
.app-header {
  display: flex; align-items: center; gap: 16px; padding: 8px 16px; min-height: 48px;
  background: var(--brand); color: #fff;
}
.brand { display: flex; align-items: center; gap: 8px; color: inherit; text-decoration: none; font-weight: 600; font-size: 1.05rem; }
.mark { flex: none; }
.mark .page { fill: #fff; }
.mark .fold { fill: rgb(255 255 255 / 60%); }
.mark .beam { fill: none; stroke: var(--brand); stroke-width: 4.5; stroke-linecap: round; stroke-linejoin: round; }
.nav { margin-left: auto; }
.nav a { color: #fff; opacity: .9; text-decoration: none; }
.nav a:hover { opacity: 1; text-decoration: underline; }
</style>
