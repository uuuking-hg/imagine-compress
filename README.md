[README.md](https://github.com/user-attachments/files/26486741/README.md)
# Light Image Compressor
<img width="1161" height="684" alt="屏幕截图 2026-04-05 151359" src="https://github.com/user-attachments/assets/e6c03e0d-3792-45e6-843f-ff4e87a5b977" />

轻量的 Windows 图片压缩工具。

用于将图片压缩到指定大小（KB）。程序优先通过调整 JPEG 质量控制体积；如果仍无法满足目标大小，则进一步缩小分辨率。

## 功能

- 图形界面
- 支持拖拽导入图片
- 支持指定目标大小（KB）
- 界面语言跟随系统，自动切换中英文
- 单文件实现，无第三方依赖

## 支持格式

输入：
`jpg` `jpeg` `png` `bmp` `gif` `tif` `tiff`

输出：
`jpg`

## 使用方式

1. 选择图片，或直接拖入窗口
2. 输入目标大小（KB）
3. 确认输出路径
4. 点击压缩

## 构建

```bash
g++ transform.cpp -std=c++17 -municode -mwindows -Wall -Wextra -lole32 -loleaut32 -lwindowscodecs -lcomdlg32 -lshell32 -o image_compressor.exe
```

## 说明

- 输出格式固定为 JPEG
- 透明区域会自动铺白底
- 目标大小过小时，可能无法达到指定值
- 核心代码位于 `transform.cpp`
