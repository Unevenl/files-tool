import utils from "./utils";
const fs = require("fs");
const xlsx = require("xlsx");
const readline = require("readline");
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});
let filesName = "";
rl.question("请输入要解析的文件名称...", (answer: string) => {
  utils.success(`文件名为：${answer},开始解析...`);
  if (answer) {
    filesName = answer;
    if (!filesName || filesName.indexOf("xlsx") === -1) {
      utils.error("非法文件名!请重试！");
      process.exit(1);
    }
    const workbook = xlsx.readFile(__dirname + "/" + filesName);
    if (!workbook) {
      utils.error("文件解析错误!");
      process.exit(1);
    }
    const sheetNames = workbook.SheetNames;
    // 获取第一个workSheet
    const sheet1 = workbook.Sheets[sheetNames[0]];
    // console.log(sheet1);

    const range = xlsx.utils.decode_range(sheet1["!ref"]);
    // console.log(range);
    interface config {
      label: string;
      cell: { c: string; r: number | string };
    }
    // columns
    const res: Array<config> = [];
    let columnsStr = "";
    for (let i = range.s.c; i < range.e.c; i++) {
      const address = { c: i, r: 0 };
      const cell: string = xlsx.utils.encode_cell(address);
      if (sheet1[cell]) {
        columnsStr += `${sheet1[cell]},`;
        res.push({ label: sheet1[cell].v, cell: { c: i, r: 0 } });
      }
    }
    utils.success(`当前表头为${columnsStr}`);
    rl.question("请输入要作为分类的依据（如“物种”):", (e: string) => {
      if (!e) {
        utils.error("字符错误!");
        process.exit(1);
      } else {
      }
      rl.close();
    });
  } else {
    utils.error("程序已退出...");
    // 实例完成
    rl.close();
  }
});

// for (let R = range.s.r; R <= range.e.r; ++R) {
//   let row_value = "";
//   for (let C = range.s.c; C <= range.e.c; ++C) {
//     let cell_address = { c: C, r: R }; //获取单元格地址
//     let cell = xlsx.utils.encode_cell(cell_address); //根据单元格地址获取单元格
//     //获取单元格值
//     if (sheet1[cell]) {
//       // 如果出现乱码可以使用iconv-lite进行转码
//       // row_value += iconv.decode(sheet1[cell].v, 'gbk') + ", ";
//       //   console.log(sheet1[cell]);
//       row_value += sheet1[cell].v + ", ";
//     }
//   }
//   console.log(row_value);
// }
