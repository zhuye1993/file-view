import { defaultOptions, renderAsync } from "docx-preview";
import renderPptx from "@/vendors/pptx";
import renderSheet from "@/vendors/xlsx";
import renderPdf from "@/vendors/pdf";
import renderImage from "@/vendors/image";
import renderText from "@/vendors/text";
import renderMp4 from "@/vendors/mp4";

// 假装构造一个vue的包装，让上层统一处理销毁和替换节点
const VueWrapper = (el) => ({
  $el: el,
  $destroy() {
    // 什么也不需要 nothing to do
  },
});

const handlers = [
  // 使用docxjs支持，目前效果最好的渲染器
  {
    accepts: ["docx"],
    handler: async (buffer, target) => {
      const docxOptions = Object.assign(defaultOptions, {
        debug: true,
        experimental: true,
      });
      await renderAsync(buffer, target, null, docxOptions);
      return VueWrapper(target);
    },
  },
  // 使用pptx2html，已通过默认值更替
  {
    accepts: ["pptx"],
    handler: async (buffer, target) => {
      await renderPptx(buffer, target, null);
      window.dispatchEvent(new Event("resize"));
      return VueWrapper(target);
    },
  },
  // 使用sheetjs + handsontable，无样式
  {
    accepts: ["xlsx"],
    handler: async (buffer, target) => {
      return renderSheet(buffer, target);
    },
  },
  // 使用pdfjs，渲染pdf，效果最好
  {
    accepts: ["pdf"],
    handler: async (buffer, target) => {
      return renderPdf(buffer, target);
    },
  },
  // 图片过滤器
  {
    accepts: ["gif", "jpg", "jpeg", "bmp", "tiff", "tif", "png", "svg"],
    handler: async (buffer, target) => {
      return renderImage(buffer, target);
    },
  },
  // 纯文本预览
  {
    accepts: [
      "txt",
      "json",
      "js",
      "css",
      "java",
      "py",
      "html",
      "jsx",
      "ts",
      "tsx",
      "xml",
      "md",
      "log",
    ],
    handler: async (buffer, target) => {
      return renderText(buffer, target);
    },
  },
  // 视频预览，仅支持MP4
  {
    accepts: ["mp4"],
    handler: async (buffer, target) => {
      renderMp4(buffer, target);
      return VueWrapper(target);
    },
  },
  // 错误处理
  {
    accepts: ["error"],
    handler: async (buffer, target, type) => {
      target.innerHTML = `<div style="text-align: center; margin-top: 80px">不支持.${type}格式的在线预览，请下载后预览或转换为支持的格式</div>
<div style="text-align: center">支持docx, xlsx, pptx, pdf, 以及纯文本格式和各种图片格式的在线预览</div>`;
      return VueWrapper(target);
    },
  },
];

// 匹配
export default handlers.reduce((result, { accepts, handler }) => {
  accepts.forEach((type) => (result[type] = handler));
  return result;
}, {});
