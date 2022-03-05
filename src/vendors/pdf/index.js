import Vue from 'vue';
import PdfView from './PdfView';

export default async function renderPdf(buffer, target) {
  return new Vue({
    render: h => h(PdfView, { props: { data: buffer } }),
  }).$mount(target)
}
