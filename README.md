# 🎯 课堂随机点名器

[![Deploy Pages](https://github.com/openclawzhangchong/ketang-dianmingqi/actions/workflows/pages.yml/badge.svg)](https://github.com/openclawzhangchong/ketang-dianmingqi/actions/workflows/pages.yml)

📄 **文档**：[开发手册](docs/DEVELOPMENT.md)（架构 / 数据契约 / 测试 / 二次开发指南）· [更新日志](CHANGELOG.md)

为计算机机房/教室场景做的随机点名工具，**两种形态**：

| 形态 | 适合场景 | 入口 |
|---|---|---|
| **网页版** | 任何设备打开即用，零安装 | **[GitHub Pages 在线版](https://openclawzhangchong.github.io/ketang-dianmingqi/)**（[备用镜像](https://mineral-donkey-ggs3.here.now/)） |
| **桌面版**（教师机） | 机房教师机常驻、大屏投影、不依赖浏览器缓存 | `desktop/启动课堂点名器.bat` 双击 |

## 界面预览

![点名主界面](docs/screenshots/home.png)

![点名结果](docs/screenshots/result.png)

## 功能

- **导入花名册**：支持 `.xlsx / .xls / .csv`（学籍系统导出的大表也能直接吃），自动识别「姓名」「班级」列，
  排除证件号/电话/父母姓名等敏感列；识别错了可手动换列
- **随机点名**：大字滚动动画，按 `空格键` 或点按钮
- **不重复模式**：一人一堂课只点一次，「重置本轮」清零
- **小组抽取**：2~6 人小组随机
- **按班级过滤**：多个班轮换上课时，只在当前班里抽
- **点名记录导出**：TXT（含时间戳），可留存
- **同名学生支持**：同名同班的不同学生会各自独立（按整行去重，不误伤）

## 隐私

- 网页版：名单只保存在**本机浏览器 localStorage**，不上传任何服务器
- 桌面版：名单只存在本机 `server/data/`
- 仓库不含任何真实学生数据（`示例花名册.xlsx` 为虚构名单）

## 桌面版使用

1. 双击 `desktop/启动课堂点名器.bat`（需本机装有 Python 3.10+，Windows 下自动开浏览器）
2. 导入花名册 → 点名。数据存本机，离线可用。

技术栈：桌面版 Python/Flask + 原生前端；网页版纯前端（SheetJS 解析表格）。

## License

MIT © 2026 智策ZCagent
