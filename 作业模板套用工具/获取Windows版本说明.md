# 获取 Windows 版本可执行文件的步骤

## 方法一：GitHub Actions 自动构建（推荐）

1. 将整个项目文件夹上传到 GitHub 仓库
2. 在 GitHub 仓库页面，点击 "Actions" 标签
3. 选择 "build-windows-exe" 工作流
4. 点击 "Run workflow" 按钮
5. 等待构建完成（通常 3-5 分钟）
6. 在完成的构建中下载 "zhuxin-word-template-exe" 文件
7. 解压后得到：`竹心Word作业模板批量套用.exe`

## 方法二：在 Windows 电脑上本地构建

如果你有 Windows 电脑，可以：

1. 安装 Python 3.9+（从 python.org 下载）
2. 将项目复制到 Windows 电脑
3. 双击运行 `build_windows.bat`
4. 在 `dist` 文件夹找到生成的 .exe

## 方法三：使用在线服务

一些在线服务可以帮你将 Python 代码转换为 Windows 可执行文件，但需要手动操作。

## 推荐

建议使用方法一（GitHub Actions），最简单且无需本地环境。
