<template>
  <div>
    <div>
      <hot-table ref="table" :settings="hotSettings"></hot-table>
    </div>
    <div class="btn-group">
      <button
        v-for="sheet in sheets"
        :key="sheet.id"
        style="padding: 0 30px"
        :type="sheetIndex === sheet.id ? 'primary' : 'default'"
        @click="handleSheet(sheet.id)"
      >
        {{ sheet.name }}
      </button>
    </div>
  </div>
</template>

<script>
import { HotTable } from "@handsontable/vue";
import Handsontable from "handsontable";
import { registerLanguageDictionary, zhCN } from "handsontable/i18n";
import { indexedColors } from "./color";
import { camelCase, captain, fixMatrix } from "./util";

// 注册中文
registerLanguageDictionary(zhCN);

// 边框类型
const borders = ["left", "right", "top", "bottom"];

export default {
  name: "HelloWorld",
  props: {
    msg: String,
    workbook: Object,
  },
  components: { HotTable },
  data() {
    return {
      sheetIndex: 0,
      selection: {
        style: {},
        ranges: [],
      },
    };
  },
  created() {
    // 注册自定义渲染
    Handsontable.renderers.registerRenderer(
      "styleRender",
      (hotInstance, TD, row, col, prop, value, cell) => {
        Handsontable.renderers.getRenderer("text")(
          hotInstance,
          TD,
          row,
          col,
          prop,
          value,
          cell
        );
        if (this.ws && cell.style) {
          const {
            style: { alignment: { wrapText } = {}, border, fill, font },
          } = cell;
          const style = TD.style;
          if (font) {
            if (font.bold) style.fontWeight = "bold";
            if (font.size) style.fontSize = `${font.size}px`;
          }
          if (fill) {
            if (fill.bgColor) {
              const { argb, indexed } = fill.bgColor;
              style.backgroundColor = `#${argb || indexedColors[indexed]}`;
            }
            if (fill.fgColor) {
              const { theme, indexed } = fill.fgColor;
              if (theme && this.themeColors) {
                const color = this.themeColors[theme + 1];
                if (color) {
                  style.color = `#${color}`;
                }
              }
              if (indexed) {
                style.color = `#${indexedColors[indexed]}`;
              }
            }
          }
          if (border) {
            borders
              .map((key) => ({ key, value: border[key] }))
              .filter((v) => v.value)
              .forEach((v) => {
                const {
                  key,
                  value: { style: borderStyle },
                } = v;
                const prefix = `border${captain(key)}`;
                if (borderStyle === "thin") {
                  style[`${prefix}Width`] = "1px";
                } else {
                  style[`${prefix}Width`] = "2px";
                }
                style[`${prefix}Style`] = "solid";
                style[`${prefix}Color`] = "#000";
              });
          }
        }
        // 启用了内联css，直接赋值
        if (cell.css) {
          const style = TD.style;
          const { css } = cell;
          Object.keys(css).forEach((key) => {
            const k = camelCase(key);
            style[k] = css[key];
          });
        }
      }
    );
  },
  watch: {
    workbook() {
      this.parseTheme();
      this.updateTable();
    },
  },
  computed: {
    hotSettings() {
      return {
        language: "zh-CN",
        readOnly: true,
        data: this.data,
        cell: this.cell,
        mergeCells: this.merge,
        colHeaders: true,
        rowHeaders: true,
        height: "calc(100vh - 107px)",
        // contextMenu: true,
        // manualRowMove: true,
        // 关闭外部点击取消选中时间的行为
        outsideClickDeselects: false,
        // fillHandle: {
        //   direction: 'vertical',
        //   autoInsertRow: true
        // },
        // afterSelectionEnd: this.afterSelectionEnd,
        // bindRowsWithHeaders: 'strict',
        licenseKey: "non-commercial-and-evaluation",
      };
    },
    ws() {
      const { workbook: { getWorksheet } = {} } = this;
      if (getWorksheet) {
        const index = this.sheetIndex || this.sheets[0].id;
        return this.workbook.getWorksheet(index);
      }
      return null;
    },
    sheets() {
      if (this.workbook.worksheets) {
        return this.workbook.worksheets.filter((sheet) => sheet._rows.length);
      }
      return [];
    },
    merge() {
      const { ws: { _merges: merges = {} } = {} } = this;
      return Object.values(merges).map(({ left, top, right, bottom }) => {
        // 构建区域
        return {
          row: top - 1,
          col: left - 1,
          rowspan: bottom - top + 1,
          colspan: right - left + 1,
        };
      });
    },
    data() {
      return fixMatrix(
        this.ws.getRows(1, this.ws.actualRowCount).map((row) =>
          row._cells.map((item) => {
            const value = item.model.value;
            if (value) {
              return value.richText ? value.richText.text : value;
            }
            return "";
          })
        ),
        this.cols.length
      );
    },
    cols() {
      return this.ws.columns.map((item) => item.letter);
    },
    columns() {
      return this.ws.columns.map((item) => ({
        ...(item.width
          ? { width: item.width < 100 ? 100 : item.width }
          : { width: 100 }),
        className: this.alignToClass(item.alignment || {}),
        renderer: "styleRender",
      }));
    },
    cell() {
      return this.ws.getRows(1, this.ws.actualRowCount).flatMap((row, ri) => {
        return row._cells
          .map((cell, ci) => {
            if (cell.style) {
              return {
                row: ri,
                col: ci,
                ...(cell.alignment
                  ? { className: this.alignToClass(cell.alignment) }
                  : {}),
                style: cell.style,
              };
            }
          })
          .filter((i) => i);
      });
    },
    border() {
      return this.ws.getRows(1, this.ws.actualRowCount).flatMap((row, ri) => {
        return row._cells
          .map((cell, ci) => {
            if (cell.style && cell.style.border) {
              const border = cell.style.border;
              const keys = Object.keys(border);
              if (keys.length) {
                return {
                  row: ri,
                  col: ci,
                  ...keys.reduce((result, key) => {
                    result[key] = {
                      width: 1,
                      color: `#${
                        (border.color && indexedColors[border.color.indexed]) ||
                        border.argb ||
                        "000000"
                      }`,
                    };
                    return result;
                  }, {}),
                };
              }
            }
          })
          .filter((i) => i);
      });
    },
  },
  methods: {
    hotTable() {
      return this.$refs.table.hotInstance;
    },
    updateTable() {
      this.hotTable().updateSettings({
        mergeCells: this.merge,
        data: this.data,
        colHeaders: this.cols,
        columns: this.columns,
        cell: this.cell,
        // customBorders: this.border,
      });
    },
    alignToClass({ horizontal, vertical }) {
      return [horizontal, vertical]
        .filter((i) => i)
        .map((key) => `ht${key.charAt(0).toUpperCase()}${key.slice(1)}`)
        .join(" ");
    },
    parseTheme() {
      const theme = this.workbook._themes.theme1;
      const parser = new DOMParser();
      if (theme) {
        const doc = parser.parseFromString(theme, "text/xml");
        const [{ children = [] } = {}] =
          doc.getElementsByTagName("a:clrScheme");
        this.themeColors = [...children]
          .flatMap((node) => [...node.getElementsByTagName("a:srgbClr")])
          .map((node) => node.getAttribute("val"))
          .filter((i) => i);
      }
    },
    // 切换sheet
    handleSheet(index) {
      if (this.sheetIndex !== index) {
        this.sheetIndex = index;
        this.$nextTick(() => {
          this.updateTable();
        });
      }
    },
    // 处理样式
    handleStyle(style, { type, key }) {
      this.selection.style = style;
      const hot = this.hotTable();
      // 暂停自定义渲染逻辑
      hot.suspendRender();
      this.selection.ranges.forEach(({ r, c }) => {
        const { css = {} } = hot.getCellMeta(r, c);
        const merged = { ...css };
        // 差量赋值，按照excel标准
        if (type === "remove") {
          delete merged[key];
        } else if (type === "add") {
          merged[key] = style[key];
        }
        hot.setCellMetaObject(r, c, {
          css: merged,
        });
      });
      // 手动渲染
      hot.render();
      // 恢复自动渲染逻辑
      hot.resumeRender();
    },
    // 选中区域回调
    afterSelectionEnd(row, column, row2, column2, selectionLayerLevel) {
      const ranges = [];
      for (let r = row; r <= row2; r++) {
        for (let c = column; c <= column2; c++) {
          ranges.push({ r, c });
        }
      }
      // 获得左上角的元数据，初始化一些状态
      const { css = {} } = this.hotTable().getCellMeta(row, column);
      this.selection.style = css;
      this.selection.ranges = ranges;
    },
  },
};
</script>

<style>
.handsontable {
  font-size: 13px;
  color: #222;
}
</style>
<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
.sheet-btn.active {
  background-color: aquamarine;
}

.btn-group {
  margin-top: 5px;
  display: block;
  border-bottom: 1px solid grey;
  background-color: lightblue;
}

.table-tool {
  padding: 8px 0;
  border-top: 1px solid black;
}
</style>
