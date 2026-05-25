// Browser stubs for Node builtins the repair path never calls. Must satisfy
// module-eval-time accesses (e.g. fs.constants.O_CREAT in `tmp`/`graceful-fs`).
const die = (n) => () => { throw new Error(`node stub: ${n} unavailable in browser`); };
export const constants = { O_CREAT: 0o100, O_EXCL: 0o200, O_RDONLY: 0, O_WRONLY: 1, O_RDWR: 2, O_TRUNC: 0o1000, O_APPEND: 0o2000, F_OK: 0, R_OK: 4, W_OK: 2, X_OK: 1, COPYFILE_EXCL: 1 };
export const existsSync = () => false;
export const readFileSync = die("readFileSync");
export const readdirSync = die("readdirSync");
export const statSync = die("statSync");
export const mkdirSync = () => {};
export const writeFileSync = die("writeFileSync");
export const rmSync = () => {};
export const openSync = die("openSync");
export const closeSync = () => {};
export const realpathSync = (p) => p;
export const promises = { readFile: die("fs.readFile"), writeFile: die("fs.writeFile"), mkdir: async () => {}, rm: async () => {}, stat: die("fs.stat"), access: die("fs.access") };
export const createRequire = () => { throw new Error("createRequire unavailable in browser"); };
export const fileURLToPath = (u) => String(u);
export const pathToFileURL = (p) => ({ href: String(p) });
export const spawnSync = die("spawnSync");
export const tmpdir = () => "/tmp";
export const platform = () => "browser";
export const EventEmitter = class { on() {} emit() {} once() {} removeListener() {} };
export const randomBytes = die("randomBytes");
const all = { constants, existsSync, readFileSync, readdirSync, statSync, mkdirSync, writeFileSync, rmSync, openSync, closeSync, realpathSync, promises, createRequire, fileURLToPath, pathToFileURL, spawnSync, tmpdir, platform, EventEmitter, randomBytes };
export default all;
