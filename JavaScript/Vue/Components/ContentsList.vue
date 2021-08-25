<template>
  <div class="contents-list">
    <div class="search">
      <input v-model="searchText" type="search" placeholder="質問タイトル検索...">
    </div>
    <div class="content" v-for="content in search" :key="content.title">
      <span>{{ content.category }}</span>
      <span>Answer: {{ content.answerCount }}</span>
      <span>Auther: {{ content.author }}</span>
      <span>Create date: {{ content.createDate }}</span>
      <h2>{{ content.title }}</h2>
    </div>
  </div>
</template>

<script>
export default {
  name: 'ContentsList',
  data() {
    return {
      searchText: '',
      contents: [
        {
          title: 'test contents',
          category: 'IT',
          answerCount: 2,
          author: 'Shou',
          createDate: "2021-01-01"
        },
      ]
    }
  },
  computed: {
    search: function() {
      return this.contents.filter(content => {
        return content.title.includes(this.searchText)
      })
    }
  }
}
</script>

<style scoped>
.search input {
  border: 1px solid rgba(0, 0, 0, .2);
  border-radius: 50px 20px 50px 20px;
  padding: 10px 20px;
  width: 500px;
  outline: none;
}

.contents-list {
  color: #666;
  margin: 100px auto;
  width: 60%;
}

.content {
  border-bottom: 1px solid rgba(0, 0, 0, .2);
  margin: 20px 0 0 0;
  padding: 20px;
  text-align: left;
  width: 100%;
  height: 100%;
}

.content h2 {
  font-size: 20px;
  margin: 10px 0 20px 0;
}

.content span {
  font-size: 16px;
  font-weight: 600;
  margin: 0 20px 0 0;
}
</style>