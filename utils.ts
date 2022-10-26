class utils {
  success(e: string) {
    console.log("\x1B[36m%s\x1B[0m", e);
  }
  error(e: string) {
    console.log("\x1B[31m%s\x1B[0m", e);
  }
}
export default new utils();
