<template>
  <iframe ref="frame" :src="src" style="min-height: calc(100vh - 185px); width:100%; border: 0" />
</template>

<script>
import axios from 'axios';
/**
 * 嵌入式预览，可支持任意网页任意情况的嵌入
 * 本实例主要展示通过axios获取二进制数据然后推送到预览页的情况
 */
export default {
  name: 'Embedded',
  data() {
    return {
      src: '',
    };
  },
  methods: {
    loadFromUrl() {
      // 要预览的文件地址
      const url = 'https://flyfish.group/%E6%95%B0%E6%8D%AE%E4%B8%AD%E5%8F%B0%E7%AC%94%E8%AE%B0(1).docx';
      // 查看器的源，当前示例为本源
      const viewerOrigin = location.origin;
      // 拼接iframe请求url
      this.src = `${viewerOrigin}?name=${encodeURIComponent(name)}&from=${encodeURIComponent(location.origin)}`;
      this.$nextTick(() => {
        const frame = this.$refs.frame;
        frame.onload = () => {
          axios({
            url,
            method: 'get',
            responseType: 'blob',
          }).then(data => {
            if (!data) {
              console.error('文件下载失败');
            }
            console.log(data)
            frame.contentWindow.postMessage(data, viewerOrigin);
          })
        }
      })
    }
  }
}
</script>

<style scoped>

</style>
