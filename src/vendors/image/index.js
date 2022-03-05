import Vue from 'vue';
import ImageViewer from './ImageViewer';
import { readDataURL } from '@/components/util';

/**
 * 图片渲染
 */
export default async function renderImage(buffer, target) {
  const url = await readDataURL(buffer);
  return new Vue({
    render: h => h(ImageViewer, { props: { image: url } }),
  }).$mount(target)
}
