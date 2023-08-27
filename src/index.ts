import { Plugin } from "siyuan";
import { getBlockAttrs, pushErrMsg, setBlockAttrs } from "./api";

export default class PluginTableCompute extends Plugin {
  private onEditorContentBindThis = this.onEditorContent.bind(this);
  private formulas = {}; //公式
  private table: HTMLTableElement;
  private blockId: BlockId;
  private tableData = {}; //值 smart data accessing object,主要依靠get获取值
  /*private cells: {
    ele: HTMLTableCellElement;
    id: string;
  }[] = []; //单元格引用*/
  private cellOnFocus: HTMLTableCellElement;
  onload() {
    console.log(this.i18n.helloPlugin);
  }

  onLayoutReady() {
    this.eventBus.on("click-editorcontent", this.onEditorContentBindThis);
  }

  onunload() {
    this.eventBus.off("click-editorcontent", this.onEditorContentBindThis);
    this.saveResult();
    console.log(this.i18n.byePlugin);
  }
  private async onEditorContent({ detail }: any) {
    //console.log(detail);
    //获取table元素和块id
    let table: HTMLTableElement;
    let blockEle = detail.event.target as HTMLElement;
    let blockId = blockEle.getAttribute("data-node-id");
    while (!blockId && blockEle) {
      blockId = blockEle.getAttribute("data-node-id");
      if (blockEle.nodeName === "TABLE") {
        table = blockEle;
      }
      blockEle = blockEle.parentElement;
    }
    //从table到其他，储存结果，结束
    if (!table) {
      this.table ? await this.saveResult() : "";
      return;
    }
    //从其他到table，开始计算
    if (!this.table) {
      await this.initCompute(table, blockId);
    } else if (
      //验证是否为同一个table，从table到table，储存结果并开始计算
      this.blockId !== blockId ||
      table.rows.length !== this.table.rows.length ||
      table.rows[0].cells.length !== this.table.rows[0].cells.length
    ) {
      await this.saveResult();
      await this.initCompute(table, blockId);
    }
    //同一个table什么都不做
  }
  private async initCompute(table: HTMLTableElement, blockId: string) {
    console.log("init");
    this.table = table;
    this.blockId = blockId;
    //设置样式
    let style = document.createElement("style");
    style.id = "PluginTableComputeStyle";
    document.getElementsByTagName("head")[0].appendChild(style);
    //let s = document.styleSheets[document.styleSheets.length - 1];
    style.innerText = `
    div[data-node-id="${blockId}"] table {
        counter-reset: rowNum;
      }
      div[data-node-id="${blockId}"] tr::before {
        counter-increment: rowNum;
        content: counter(rowNum);
        background-color: aliceblue;
        color: black;
      }
      div[data-node-id="${blockId}"] tr {
        counter-reset: colNum;
      }
      div[data-node-id="${blockId}"] tr:nth-child(1) > th::before {
        counter-increment: colNum;
        content: counter(colNum, upper-alpha);
        position: absolute;
        top: -30px;
        background-color: aliceblue;
        color: black;
      }
    `;
    //window.getSelection()
    document.addEventListener("selectionchange", this.onSelectionChange);
    //table.onfocus = this.fieldOnfocus;
    const attrs = await getBlockAttrs(this.blockId);
    const formulasOldString = attrs["custom-plugin-table-compute-formulas"];
    const formulasOld = formulasOldString ? JSON.parse(formulasOldString) : {};
    for (let row of this.table.rows) {
      for (let cell of row.cells) {
        const index = this.getABCindex(cell);
        this.formulas[index] = formulasOld[index] || cell.textContent;
        let get = () => {
          //const tableData = this.tableData;
          let value = this.formulas[index] || "";
          if ("=" == value.charAt(0)) {
            //生成计算式 evaluate the formula
            let evalString = `with(tableData){
                return ${value.substring(1)};
            }`;
            let calcFunc = Function("tableData", evalString);
            return calcFunc(this.tableData);
          } else {
            // return value as it is, convert to number if possible:
            return isNaN(parseFloat(value)) ? value : parseFloat(value);
          }
        };
        // Add smart getter to the data array for both upper and lower case variants:
        Object.defineProperty(this.tableData, index, { get });
        Object.defineProperty(this.tableData, index.toLowerCase(), { get });
      }
    }
    this.onSelectionChange(); //手动触发一次
  }
  private onSelectionChange = () => {
    let select = document.getSelection().anchorNode;
    while (select && select.nodeName != "TD" && select.nodeName != "TH") {
      select = select.parentElement;
    }
    let cell = select as HTMLTableCellElement;
    //未改变
    if (cell === this.cellOnFocus) {
      return;
    }
    //onblur,缓存公式，重新计算并显示结果
    if (this.cellOnFocus) {
      const indexOld = this.getABCindex(this.cellOnFocus);
      this.formulas[indexOld] = this.cellOnFocus.textContent;
      //重新计算
      for (let row of this.table.rows) {
        for (let cell of row.cells) {
          const index = this.getABCindex(cell);
          try {
            if (
              this.formulas[index] &&
              this.formulas[index].charAt(0) == "="
              //防止清除样式
            ) {
              cell.textContent = this.tableData[index];
            }
          } catch (e) {
            /* ignore */
            let warn = `单元格 : ${index}</br>
               ${e.name} : ${e.message}</br>
            `;
            switch (e.message) {
              case "Maximum call stack size exceeded":
                warn += "可能存在循环引用，请检查</br>";
                break;
            }
            pushErrMsg(warn);
          }
        }
      }
    }
    //onfocus，显示公式
    if (select) {
      const index = this.getABCindex(cell);
      const formula = this.formulas[index] as string;
      if (formula && formula.charAt(0) == "=") {
        //防止清除样式
        cell.textContent = formula;
      }
      this.cellOnFocus = cell;
    }
    //console.log("formulas", this.formulas);
    //console.log("tableData", this.tableData);
  };
  private async saveResult() {
    console.log("save");
    if (!this.table) {
      return;
    }
    document.getElementById("PluginTableComputeStyle")?.remove();
    this.onSelectionChange(); //手动触发一次
    await setBlockAttrs(this.blockId, {
      "custom-plugin-table-compute-formulas": JSON.stringify(this.formulas),
    });
    document.removeEventListener("selectionchange", this.onSelectionChange);
    this.tableData = {};
    this.formulas = {};
    this.table = null;
  }
  private getABCindex(cell: HTMLTableCellElement) {
    if (!cell) {
      return;
    }
    const colIndex = cell.cellIndex + 1;
    const row = cell.parentElement as HTMLTableRowElement;
    const rowIndex = row.rowIndex + 1;
    return this.num2ABC(colIndex) + rowIndex;
  }
  private num2ABC(num: number) {
    if (num <= 0) {
      return null;
    }
    let result = "";
    while (num > 26) {
      let numTamp = num % 26;
      if (numTamp === 0) {
        numTamp = 26;
        num = Math.floor(num / 26) - 1;
      } else {
        num = Math.floor(num / 26);
      }
      result = String.fromCharCode(65 + numTamp - 1) + result;
    }
    result = String.fromCharCode(65 + num - 1) + result;
    return result;
  }
  private ABC2num(cell: string) {
    let col = 0;
    let codeA = "A".charCodeAt(0);
    const chars = cell.replace(/[0-9]/g, "").toUpperCase().split("");
    for (let i = 0; i < chars.length; i++) {
      let char = chars[i];
      const code = char.charCodeAt(0);
      const minus = code - codeA + 1;
      col += minus * 26 ** (chars.length - 1 - i);
    }
    const row = parseInt(cell.replace(/[^0-9]/gi, ""));
    return { col, row };
  }
}
