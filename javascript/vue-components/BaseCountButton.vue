<template>
  <button @click="emit('increment')" class="base-count-button">increment</button>
  <button @click="emit('decrement')" class="base-count-button">decrement</button>
</template>

<script setup>
// vue@3.2.29
const emit = defineEmits(["increment", "decrement"])
</script>

<style scoped>
.base-count-button {
  display: block;
  background-color: #666;
  border: none;
  border-radius: 5px;
  color: #eee;
  width: 100%;
  height: 35px;
  cursor: pointer;
}
</style>