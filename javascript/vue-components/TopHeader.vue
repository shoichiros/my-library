<template>
  <div class="top-header">
    <header class="top-header-contents">
      <h1>{{ siteTitle }}</h1>
      <nav class="top-header-nav">
        <a v-for="list in navLists" :key="list.name" href="">{{ list.name }}</a>
      </nav>
    </header>
  </div>
</template>

<script>
export default {
  name: 'TopHeader',
  props: {
    siteTitle: String
  },
  data() {
    return {
      navLists: [
        { name: 'About' },
        { name: 'Q&A' },
        { name: 'Blog' }
      ]
    }
  }
}
</script>

<style scoped>
.top-header {
  background: rgba(255, 255, 255, .8);
  position: fixed;
  width: 100%;
  height: 80px;
}

.top-header-contents {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin: 0 auto;
  width: 90%;
  height: 100%;
}

.top-header-contents h1 {
  color: #555;
  font-size: 24px;
}

.top-header-nav {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.top-header-nav a {
  color: #555;
  font-weight: 600;
  margin: 0 20px;
  text-decoration: none;
}
</style>