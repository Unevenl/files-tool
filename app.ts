import utils from "./utils";
import config from "./config";
import Spinnies from "spinnies";
const spinner = { interval: 80, frames: ["🍇", "🍈", "🍉", "🍋"] };
const spinnies = new Spinnies({
  color: "blue",
  succeedColor: "green",
  spinner,
});
const fs = require("fs");
const xlsx = require("xlsx");
const readline = require("readline");
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});
class sheetCl {
  filesName: string;
  range: range;
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
  c_sort_f: number | string;
  c_sort_s: number | string;
  c_sort_type: string;
  c_sort_name: string;
  sheet: any;
  constructor() {
    this.filesName = "";
    this.range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
    this.rowStart = 0;
    this.rowEnd = 0;
    this.columnStart = 0;
    this.columnEnd = 0;
    this.c_sort_f = 0;
    this.c_sort_s = 0;
    this.c_sort_type = "";
    this.c_sort_name = "";
    rl.question("请输入运行密码>:", (pwd: string) => {
      if (pwd === config.pwd) {
        utils.success("validation succeeded!");
        this.init();
      }
    });
    rl.on("close", () => {
      utils.success("bye~");
      process.exit(1);
    });
  }
  init() {
    utils.success("files loading...");
    rl.question("请输入要解析的文件名称>:", (answer: string) => {
      utils.success(`文件名为：${answer}`);
      utils.success(`start running...`);
      if (!answer) {
        utils.error("program exit...");
        rl.close();
      }
      this.filesName = answer;
      if (!this.filesName || this.filesName.indexOf("xlsx") === -1) {
        utils.error("请使用xlsx类文件!请重试！");
        //process.exit(1)
      }
      let workbook;
      try {
        workbook = xlsx.readFile("./" + this.filesName);
      } catch (error) {
        utils.error(`查找文件失败!${"./" + this.filesName}`);
        //process.exit(1)
      }
      //   try {
      //     workbook = xlsx.readFile(__dirname + "/" + this.filesName);
      //   } catch (error) {
      //     utils.error(`查找文件失败!${__dirname + "/" + this.filesName}`);
      //     //process.exit(1)
      //   }

      if (!workbook) {
        utils.error("文件解析错误!");
        //process.exit(1)
      }
      const sheetNames = workbook.SheetNames;
      this.sheet = workbook.Sheets[sheetNames[0]];
      this.range = xlsx.utils.decode_range(this.sheet["!ref"]);
      this.rowStart = this.range.s.r;
      this.rowEnd = this.range.e.r;
      this.columnStart = this.range.s.c;
      this.columnEnd = this.range.e.c;
      // columns array
      const res: Array<config> = [];
      let columnsStr = "";
      for (let i = this.range.s.c; i < this.range.e.c; i++) {
        const address = { c: i, r: 0 };
        const cell: string = xlsx.utils.encode_cell(address);
        if (this.sheet[cell]) {
          columnsStr += `${this.sheet[cell].v},`;
          res.push({ label: this.sheet[cell].v, cell: { c: i, r: 0 } });
        }
      }
      // utils.success(`当前表头为>:${columnsStr}`);
      rl.question("请输入分类依据>:", (e: string) => {
        if (!e) {
          utils.error("字符错误!");
          rl.close();
          process.exit(1);
        }
        this.c_sort_type = e;
        // get sepcial columns
        const arr: Array<string> = [];
        let cellCheck: cell = { c: "", r: "" };
        res.map((item) => {
          if (item.label === e) {
            cellCheck = item.cell;
            this.c_sort_f = cellCheck.c;
          }
        });
        for (let i = this.rowStart + 1; i < this.rowEnd; i++) {
          const cellS: string = xlsx.utils.encode_cell({
            c: cellCheck.c,
            r: i,
          });
          if (this.sheet[cellS]) {
            arr.filter((item) => item === this.sheet[cellS].v).length === 0
              ? arr.push(this.sheet[cellS].v)
              : "";
          }
        }
        // const data: Array<string> = fs.readdirSync("./build");
        rl.question("请输入要输出的文件夹名称：>", (text: string) => {
          if (!text) {
            utils.error("字符错误!");
            rl.close();
            process.exit(1);
          }
          this.mkResult(text, arr);
          this.classify(res, arr, text);
        });
      });
    });
  }
  mkResult(e: string, arr: Array<string>) {
    // dev ./build/result
    fs.exists(`./${e}`, async (exists: string) => {
      if (!exists) {
        await fs.mkdirSync(`./${e}`);
        this.mkFolder(arr, e); // create special folder
      }
    });
  }
  mkFolder(data: Array<string>, text: string) {
    // dev ./build/result
    data.map(async (item) => {
      const e: string = await fs.existsSync(`./${text}/${item}`);
      if (!e) {
        await fs.mkdirSync(`./${text}/${item}`);
      }
    });
  }
  classify(res: Array<config>, arr: Array<string>, text: string) {
    rl.question("请输入分类列>:", async (e: string) => {
      if (!e) {
        utils.error("字符错误!");
        //process.exit(1)
        rl.close();
      }
      this.c_sort_name = e;
      res.map((item) => {
        if (item.label === e) {
          this.c_sort_s = item.cell.c;
        }
      });
      utils.success("文件分类中...");
      utils.success("请耐心等待，不要退出...");
      spinnies.add("spinner-1", { text: "loading..." });
      for (let i = this.rowStart + 1; i < this.rowEnd; i++) {
        const cellName = xlsx.utils.encode_cell({
          c: this.c_sort_s,
          r: i,
        });
        const cellType = xlsx.utils.encode_cell({
          c: this.c_sort_f,
          r: i,
        });
        // console.log( `./pictures/${this.sheet[cellName].v}`,`./result/${this.sheet[cellType].v}/${this.sheet[cellName].v}`)
        // dev ./build/pictures ./build/result
        const ret = await fs.copyFileSync(
          `./pictures/${this.sheet[cellName].v}`,
          `./${text}/${this.sheet[cellType].v}/${this.sheet[cellName].v}`
        );
        if (ret) {
          utils.error(`${this.sheet[cellName].v}拷贝失败!请检查文件权限!`);
        }
        if (i == this.rowEnd - 1) {
          spinnies.succeed("spinner-1", { text: "files copy done!" });
          utils.success("go and check pictures!");
          this.init()
        }
      }

      //process.exit(1)
      //rl.close();
    });
  }
}
interface cell {
  c: string | number;
  r: string | number;
}
interface config {
  label: string;
  cell: { c: number | string; r: number | string };
}
interface range {
  s: { c: number; r: number };
  e: { c: number; r: number };
}
const sheetCls = new sheetCl();
