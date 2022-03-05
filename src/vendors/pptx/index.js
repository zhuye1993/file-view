/* global $, dimple */
"use strict";

import processPptx from "./process_pptx";
import pptxStyle from "./pptx_css";
import $ from "jquery";

/**
 * @param {ArrayBuffer} pptx
 * @param {Element|String} resultElement
 * @param {Element|String} [thumbElement]
 */
const renderPptx = (pptx, resultElement, thumbElement) => {
  const $result = $(resultElement);
  const $wrapper = $('<div class="pptx-wrapper"></div>');
  $result.html("");
  $result.append($wrapper);
  $wrapper.append(`<style>${pptxStyle}</style>`);
  let isDone = false;

  return new Promise((resolve, reject) => {
    const processMessage = (msg) => {
      if (isDone) return;
      switch (msg.type) {
        case "slide":
          $wrapper.append(msg.data);
          break;
        case "pptx-thumb":
          if (thumbElement)
            $(thumbElement).attr("src", `data:image/jpeg;base64,${msg.data}`);
          break;
        case "slideSize":
          break;
        case "globalCSS":
          $wrapper.append(`<style>${msg.data}</style>`);
          break;
        case "Done":
          isDone = true;
          processCharts(msg.data.charts);
          resolve(msg.data.time);
          break;
        case "WARN":
          console.warn("PPTX processing warning: ", msg.data);
          break;
        case "ERROR":
          isDone = true;
          console.error("PPTX processing error: ", msg.data);
          reject(new Error(msg.data));
          break;
        case "DEBUG":
          // console.debug('Worker: ', msg.data);
          break;
        case "INFO":
        default:
        // console.info('Worker: ', msg.data);
      }
    };
    /*
    // Actual Web Worker - If you want to use this, switching worker's url to Blob is probably better
    const worker = new Worker('./dist/worker.js')
    worker.addEventListener('message', event => processMessage(event.data), false)
    const stopWorker = setInterval(() => { // Maybe this should be done in the message processing
      if (isDone) {
        worker.terminate()
        // console.log("worker terminated");
        clearInterval(stopWorker)
      }
    }, 500)
    */
    const worker = {
      // shim worker
      postMessage: () => {},
      terminate: () => {},
    };
    // processMessage方法，成功之后执行
    processPptx((func) => {
      worker.postMessage = func;
    }, processMessage);
    // 执行postMessage方法
    worker.postMessage({
      type: "processPPTX",
      data: pptx,
    });
  }).then((time) => {
    const resize = () => {
      const slidesWidth = Math.max(
        ...Array.from($wrapper.children("section")).map((s) => s.offsetWidth)
      );
      const wrapperWidth = $wrapper[0].offsetWidth;
      $wrapper.css({
        transform: `scale(${wrapperWidth / slidesWidth})`,
        "transform-origin": "top left",
      });
    };
    resize();
    window.addEventListener("resize", resize);
    setNumericBullets($(".block"));
    setNumericBullets($("table td"));
    return time;
  });
};

export default renderPptx;

function processCharts(queue) {
  for (let i = 0; i < queue.length; i++) {
    processSingleChart(queue[i].data);
  }
}

function convertChartData(chartData) {
  const data = [];
  const xLabels = [];
  const groupLabels = [];
  chartData.forEach((group, i) => {
    const groupName = group.key;
    groupLabels[i] = group.key;
    group.values.forEach((value, j) => {
      const labelName = group.xlabels[j];
      xLabels[j] = group.xlabels[j];
      data.push({ name: labelName, group: groupName, value: value.y });
    });
  });
  // console.log('TRANSFORMED DATA:', (data))
  return { data, xLabels, groupLabels };
}

function processSingleChart(d) {
  const chartID = d.chartID;
  const chartType = d.chartType;
  const chartData = d.chartData;
  // console.log(`WRITING GRAPH OF TYPE ${chartType} TO ID #${chartID}:`, chartData)

  let data = [];

  switch (chartType) {
    case "lineChart": {
      const { data: data_, xLabels, groupLabels } = convertChartData(chartData);
      data = data_;
      const container = document.getElementById(chartID);
      const svg = dimple.newSvg(
        `#${chartID}`,
        container.style.width,
        container.style.height
      );

      // eslint-disable-next-line new-cap
      const myChart = new dimple.chart(svg, data);
      const xAxis = myChart.addCategoryAxis("x", "name");
      xAxis.addOrderRule(xLabels);
      xAxis.addGroupOrderRule(groupLabels);
      xAxis.title = null;
      const yAxis = myChart.addMeasureAxis("y", "value");
      yAxis.title = null;
      myChart.addSeries("group", dimple.plot.line);
      myChart.addLegend(60, 10, 500, 20, "right");
      myChart.draw();

      break;
    }
    case "barChart": {
      const { data: data_, xLabels, groupLabels } = convertChartData(chartData);
      data = data_;
      const container = document.getElementById(chartID);
      const svg = dimple.newSvg(
        "#" + chartID,
        container.style.width,
        container.style.height
      );

      // eslint-disable-next-line new-cap
      const myChart = new dimple.chart(svg, data);
      const xAxis = myChart.addCategoryAxis("x", ["name", "group"]);
      xAxis.addOrderRule(xLabels);
      xAxis.addGroupOrderRule(groupLabels);
      xAxis.title = null;
      const yAxis = myChart.addMeasureAxis("y", "value");
      yAxis.title = null;
      myChart.addSeries("group", dimple.plot.bar);
      myChart.addLegend(60, 10, 500, 20, "right");
      myChart.draw();
      break;
    }
    case "pieChart":
    case "pie3DChart": {
      // data = chartData[0].values
      // chart = nv.models.pieChart()
      // nvDraw(chart, data)
      const { data: data_, groupLabels } = convertChartData(chartData);
      data = data_;
      const container = document.getElementById(chartID);
      const svg = dimple.newSvg(
        `#${chartID}`,
        container.style.width,
        container.style.height
      );

      // eslint-disable-next-line new-cap
      const myChart = new dimple.chart(svg, data);
      const pieAxis = myChart.addMeasureAxis("p", "value");
      pieAxis.addOrderRule(groupLabels);
      myChart.addSeries("name", dimple.plot.pie);
      myChart.addLegend(50, 20, 400, 300, "left");
      myChart.draw();
      break;
    }
    case "areaChart": {
      const { data: data_, xLabels, groupLabels } = convertChartData(chartData);
      data = data_;
      const container = document.getElementById(chartID);
      const svg = dimple.newSvg(
        "#" + chartID,
        container.style.width,
        container.style.height
      );

      // eslint-disable-next-line new-cap
      const myChart = new dimple.chart(svg, data);
      const xAxis = myChart.addCategoryAxis("x", "name");
      xAxis.addOrderRule(xLabels);
      xAxis.addGroupOrderRule(groupLabels);
      xAxis.title = null;
      const yAxis = myChart.addMeasureAxis("y", "value");
      yAxis.title = null;
      myChart.addSeries("group", dimple.plot.area);
      myChart.addLegend(60, 10, 500, 20, "right");
      myChart.draw();

      break;
    }
    case "scatterChart": {
      for (let i = 0; i < chartData.length; i++) {
        const arr = [];
        for (let j = 0; j < chartData[i].length; j++) {
          arr.push({ x: j, y: chartData[i][j] });
        }
        data.push({ key: "data" + (i + 1), values: arr });
      }

      // data = chartData;
      // chart = nv.models.scatterChart()
      //   .showDistX(true)
      //   .showDistY(true)
      //   .color(d3.scale.category10().range())
      // chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'))
      // chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'))
      // nvDraw(chart, data)
      break;
    }
    default:
  }
}

function setNumericBullets(elem) {
  const paragraphsArray = elem;
  for (let i = 0; i < paragraphsArray.length; i++) {
    const buSpan = $(paragraphsArray[i]).find(".numeric-bullet-style");
    if (buSpan.length > 0) {
      // console.log("DIV-"+i+":");
      let prevBultTyp = "";
      let prevBultLvl = "";
      let buletIndex = 0;
      const tmpArry = [];
      let tmpArryIndx = 0;
      const buletTypSrry = [];
      for (let j = 0; j < buSpan.length; j++) {
        const bulletType = $(buSpan[j]).data("bulltname");
        const bulletLvl = $(buSpan[j]).data("bulltlvl");
        // console.log(j+" - "+bult_typ+" lvl: "+bult_lvl );
        if (buletIndex === 0) {
          prevBultTyp = bulletType;
          prevBultLvl = bulletLvl;
          tmpArry[tmpArryIndx] = buletIndex;
          buletTypSrry[tmpArryIndx] = bulletType;
          buletIndex++;
        } else {
          if (bulletType === prevBultTyp && bulletLvl === prevBultLvl) {
            prevBultTyp = bulletType;
            prevBultLvl = bulletLvl;
            buletIndex++;
            tmpArry[tmpArryIndx] = buletIndex;
            buletTypSrry[tmpArryIndx] = bulletType;
          } else if (bulletType !== prevBultTyp && bulletLvl === prevBultLvl) {
            prevBultTyp = bulletType;
            prevBultLvl = bulletLvl;
            tmpArryIndx++;
            tmpArry[tmpArryIndx] = buletIndex;
            buletTypSrry[tmpArryIndx] = bulletType;
            buletIndex = 1;
          } else if (
            bulletType !== prevBultTyp &&
            Number(bulletLvl) > Number(prevBultLvl)
          ) {
            prevBultTyp = bulletType;
            prevBultLvl = bulletLvl;
            tmpArryIndx++;
            tmpArry[tmpArryIndx] = buletIndex;
            buletTypSrry[tmpArryIndx] = bulletType;
            buletIndex = 1;
          } else if (
            bulletType !== prevBultTyp &&
            Number(bulletLvl) < Number(prevBultLvl)
          ) {
            prevBultTyp = bulletType;
            prevBultLvl = bulletLvl;
            tmpArryIndx--;
            buletIndex = tmpArry[tmpArryIndx] + 1;
          }
        }
        // console.log(buletTypSrry[tmpArryIndx]+" - "+buletIndex);
        const numIdx = getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
        $(buSpan[j]).html(numIdx);
      }
    }
  }
}

function getNumTypeNum(numTyp, num) {
  let rtrnNum = "";
  switch (numTyp) {
    case "arabicPeriod":
      rtrnNum = num + ". ";
      break;
    case "arabicParenR":
      rtrnNum = num + ") ";
      break;
    case "alphaLcParenR":
      rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
      break;
    case "alphaLcPeriod":
      rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
      break;

    case "alphaUcParenR":
      rtrnNum = alphaNumeric(num, "upperCase") + ") ";
      break;
    case "alphaUcPeriod":
      rtrnNum = alphaNumeric(num, "upperCase") + ". ";
      break;

    case "romanUcPeriod":
      rtrnNum = romanize(num) + ". ";
      break;
    case "romanLcParenR":
      rtrnNum = romanize(num) + ") ";
      break;
    case "hebrew2Minus":
      rtrnNum = hebrew2Minus.format(num) + "-";
      break;
    default:
      rtrnNum = num;
  }
  return rtrnNum;
}

function romanize(num) {
  if (!+num) return false;
  const digits = String(+num).split("");
  const key = [
    "",
    "C",
    "CC",
    "CCC",
    "CD",
    "D",
    "DC",
    "DCC",
    "DCCC",
    "CM",
    "",
    "X",
    "XX",
    "XXX",
    "XL",
    "L",
    "LX",
    "LXX",
    "LXXX",
    "XC",
    "",
    "I",
    "II",
    "III",
    "IV",
    "V",
    "VI",
    "VII",
    "VIII",
    "IX",
  ];
  let roman = "";
  let i = 3;
  while (i--) roman = (key[+digits.pop() + i * 10] || "") + roman;
  return new Array(+digits.join("") + 1).join("M") + roman;
}

const hebrew2Minus = archaicNumbers([
  [1000, ""],
  [400, "ת"],
  [300, "ש"],
  [200, "ר"],
  [100, "ק"],
  [90, "צ"],
  [80, "פ"],
  [70, "ע"],
  [60, "ס"],
  [50, "נ"],
  [40, "מ"],
  [30, "ל"],
  [20, "כ"],
  [10, "י"],
  [9, "ט"],
  [8, "ח"],
  [7, "ז"],
  [6, "ו"],
  [5, "ה"],
  [4, "ד"],
  [3, "ג"],
  [2, "ב"],
  [1, "א"],
  [/יה/, "ט״ו"],
  [/יו/, "ט״ז"],
  [/([א-ת])([א-ת])$/, "$1״$2"],
  [/^([א-ת])$/, "$1׳"],
]);

function archaicNumbers(arr) {
  // const arrParse = arr.slice().sort(function (a, b) { return b[1].length - a[1].length })
  return {
    format: function (n) {
      let ret = "";
      $.each(arr, function () {
        const num = this[0];
        if (parseInt(num) > 0) {
          for (; n >= num; n -= num) ret += this[1];
        } else {
          ret = ret.replace(num, this[1]);
        }
      });
      return ret;
    },
  };
}

function alphaNumeric(num, upperLower) {
  num = Number(num) - 1;
  let aNum = "";
  if (upperLower === "upperCase") {
    aNum = (
      (num / 26 >= 1 ? String.fromCharCode(num / 26 + 64) : "") +
      String.fromCharCode((num % 26) + 65)
    ).toUpperCase();
  } else if (upperLower === "lowerCase") {
    aNum = (
      (num / 26 >= 1 ? String.fromCharCode(num / 26 + 64) : "") +
      String.fromCharCode((num % 26) + 65)
    ).toLowerCase();
  }
  return aNum;
}
