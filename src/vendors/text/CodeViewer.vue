<template>
  <pre class="code-area">
    {{ value }}
  </pre>
</template>

<script>
export default {
  name: 'CodeEditor',
  props: {
    value: {
      type: String,
      description: '值'
    }
  },
};
</script>

<style scoped>
.code-area {
  display: block;
  margin: 0 auto;
  font-size: 12px;
  width: 1000px;
  min-height: 500px;
  background: #1f1f1f;
  word-break: break-word;
  white-space: break-spaces;
  color: #5af117
}
</style>
