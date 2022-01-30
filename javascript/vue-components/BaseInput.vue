<template>
  <input
    type="text"
    class="main-base-input"
    :placeholder="placeholder"
    :value="modelValue"
    @input="emit('update:modelValue', $event.target.value)"
    @change="emit('update:log')"
  />
</template>

<script setup>
defineProps({
  modelValue: {
    type: String,
    required: true
  },
  placeholder: {
    type: String,
    required: false
  }
})

const emit = defineEmits(["update:modelValue", "update:log"])
</script>

<style scoped>
.main-base-input {
  display: block;
  border: 1px solid #aaa;
  box-sizing: border-box;
  padding: 10px 8px;
  width: 200px;
}
</style>