[README.md](https://github.com/user-attachments/files/26486722/README.md)
# Light Image Compressor
<img width="1161" height="684" alt="屏幕截图 2026-04-05 151359" src="https://github.com/user-attachments/assets/c1573f18-4cc1-4a58-b58d-f269b4c0d46f" />

一个轻量的 Windows 图片压缩工具。

选图，填目标大小（KB），导出 JPEG。程序会先尽量通过 JPEG 质量控制体积，不够的话再适当缩小分辨率，目标是把大小压下来，同时别把画质弄得太难看。

## 功能

- 图形界面
- 支持拖拽图片
- 可指定目标大小（KB）
- 自动中英文界面（跟随系统语言）
- 单文件实现，无第三方依赖

## 支持格式

输入：
`jpg` `jpeg` `png` `bmp` `gif` `tif` `tiff`

输出：
`jpg`

## 使用

1. 选择图片，或者直接拖进窗口
2. 输入目标大小（KB）
3. 确认输出路径
4. 点击压缩

## 构建

```bash
g++ transform.cpp -std=c++17 -municode -mwindows -Wall -Wextra -lole32 -loleaut32 -lwindowscodecs -lcomdlg32 -lshell32 -o image_compressor.exe
```

## 说明

- 输出固定为 JPEG
- 透明区域会自动铺白底
- 如果目标太小，可能压不到
- 核心代码在 `transform.cpp`
