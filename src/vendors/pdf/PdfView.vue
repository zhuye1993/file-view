<template>
  <div class="home_wrap">
    <div class="pdf_down">
      <div class="pdf_set_left" @click="scaleD()">➕</div>
      <div class="pdf_set_middle" @click="scaleX()">➖</div>
      <!-- <div class="pdf-pre" @click="prePage">上一页</div> -->
      <!-- <div class="pdf-next" @click="nextPage">下一页</div> -->
    </div>
    <div :style="{ width: pdf_div_width, margin: '0 auto' }">
      <!-- <canvas id="the_canvas"></canvas> -->
      <canvas
        v-for="page in pdf_pages"
        :id="'the_canvas' + page"
        :key="page"
      ></canvas>
    </div>
  </div>
</template>

<script>
import * as PDFJS from "pdfjs-dist/legacy/build/pdf";
// const PDFJS = { GlobalWorkerOptions: {} };

PDFJS.GlobalWorkerOptions.workerSrc = require("pdfjs-dist/legacy/build/pdf.worker.entry.js");

export default {
  name: "PdfView",
  props: {
    data: ArrayBuffer,
  },
  data() {
    return {
      pdf_scale: 1.0, //pdf放大系数
      pdf_pages: [],
      pdf_div_width: "",
      currentPage: 1,
    };
  },
  mounted() {
    this.loadFile();
  },
  methods: {
    scaleD() {
      //放大
      let max = 0;
      if (window.screen.width > 1440) {
        max = 1.4;
      } else {
        max = 1.2;
      }
      if (this.pdf_scale >= max) {
        return;
      }
      this.pdf_scale = this.pdf_scale + 0.1;
      this.loadFile();
    },
    scaleX() {
      //缩小
      let min = 1.0;
      if (this.pdf_scale <= min) {
        return;
      }
      this.pdf_scale = this.pdf_scale - 0.1;
      this.loadFile();
    },
    async loadFile() {
      //初始化pdf
      console.log("初始化pdf");
      this.pdfDoc = await PDFJS.getDocument(this.data).promise;
      this.pdf_pages = this.pdfDoc.numPages;
      this.$nextTick(() => this.renderPage());
    },
    async renderPage(num = 1) {
      this.currentPage = num;
      //渲染pdf页
      console.log("渲染pdf页");
      const page = await this.pdfDoc.getPage(num);
      const canvas = document.getElementById("the_canvas" + num);
      // const canvas = document.getElementById("the_canvas");
      const ctx = canvas.getContext("2d");
      const dpr = window.devicePixelRatio || 1;
      const bsr =
        ctx.webkitBackingStorePixelRatio ||
        ctx.mozBackingStorePixelRatio ||
        ctx.msBackingStorePixelRatio ||
        ctx.oBackingStorePixelRatio ||
        ctx.backingStorePixelRatio ||
        1;
      const ratio = dpr / bsr;
      const viewport = page.getViewport({ scale: this.pdf_scale });
      canvas.width = viewport.width * ratio;
      canvas.height = viewport.height * ratio;
      canvas.style.width = viewport.width + "px";
      this.pdf_div_width = viewport.width + "px";
      canvas.style.height = viewport.height + "px";
      ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
      const renderContext = {
        canvasContext: ctx,
        viewport: viewport,
      };
      page.render(renderContext);
      if (this.pdf_pages > num) {
        setTimeout(() => {
          return this.renderPage(num + 1);
        });
      }
    },
    nextPage() {
      if (this.pdf_pages > this.currentPage) {
        this.renderPage(this.currentPage + 1);
      }
    },
    prePage() {
      if (this.currentPage > 1) {
        this.renderPage(this.currentPage - 1);
      }
    },
  },
};
</script>

<style scoped>
.home_wrap {
  width: 100%;
  height: 100%;
}
.home_wrap .pdf_down {
  position: fixed;
  display: flex;
  z-index: 20;
  right: 26px;
  bottom: 7%;
  cursor: pointer;
}
.home_wrap .pdf-pre {
  position: fixed;
  display: flex;
  z-index: 20;
  right: 160px;
  bottom: 9%;
  cursor: pointer;
}
.home_wrap .pdf-next {
  position: fixed;
  display: flex;
  z-index: 20;
  right: 100px;
  bottom: 9%;
}
.home_wrap .pdf_down .pdf_set_left {
  width: 30px;
  height: 40px;
  color: #408fff;
  font-size: 15px;
  padding-top: 25px;
  text-align: center;
  margin-right: 5px;
  cursor: pointer;
}
.home_wrap .pdf_down .pdf_set_middle {
  width: 30px;
  height: 40px;
  color: #408fff;
  font-size: 15px;
  padding-top: 25px;
  text-align: center;
  margin-right: 5px;
  cursor: pointer;
}
</style>
