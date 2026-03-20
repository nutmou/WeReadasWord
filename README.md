# 📚 微信读书 - Word 2019 专业增强风格 (v1.0.0 完美摸鱼版)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

这是一个将微信读书网页版转换为 **Microsoft Word 2019 专业增强版（Professional Plus）风格**的 Tampermonkey 脚本。不仅提供舒适的沉浸式阅读体验，更内置了多重硬核的“反老板”机制，是专为职场人量身定制的终极摸鱼阅读神器。

---

## ✨ 核心功能特点

- **像素级拟真界面**：完美复刻 Word 2019 专版特征（居中标题、蓝色选项卡、操作说明搜索）。
- **沉浸式 UI 组件**：内置左侧导航窗格（带真实空结果提示语）、高拟真动态双侧数字标尺（带缩进游标）、底部状态栏（字数统计、视图、缩放比例）以及右侧仿真滚动条。
- **🛡️ 老板键（一键白纸）**：点击右上角“最小化（`-`）”按钮，可瞬间隐藏小说正文并禁用鼠标点击防误触，屏幕只留下带有标尺和界面的 Word 空白框架。点击“最大化（`□`）”即可无缝恢复阅读。
- **⚡ 全局快捷键启停**：随时按下键盘 <kbd>Alt</kbd> + <kbd>W</kbd>，即可在一秒内彻底关闭或重新开启伪装，在“原版微信读书”与“Word 伪装版”之间丝滑切换，应对突发查岗。
- **智能翻页排版**：自动计算并调整阅读区域边距，悬浮翻页按钮智能避让导航栏与滚动条，保障阅读顺畅不遮挡。
- **纯本地零依赖**：所有界面图标均采用纯代码 SVG 绘制，无外部图片请求，加载极快且稳定。

---

## 📦 安装指南

1. **首先确保你的浏览器已安装 Tampermonkey（油猴）扩展**：
   - [Chrome 商店链接](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
   - [Firefox 附加组件链接](https://addons.mozilla.org/zh-CN/firefox/addon/tampermonkey/)
   - [Edge 插件链接](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)

2. **安装本脚本**：
   - **方法 1**：直接点击 [[WeReadFaker_Organized/weread-word-style-final-v1.0.0.user.js](https://github.com/nutmou/WeReadasWord/blob/f722bf8e17be6d6d967e280a3775227a949df328/WeReadFaker_Organized/weread-word-style-final-v1.0.0.user.js)] 
   - **方法 2**：复制其中的完整代码，在 Tampermonkey 的管理面板中点击“添加新脚本”，粘贴并保存。

3. **确保脚本已启用**：
   - 在浏览器的 Tampermonkey 插件菜单中，确认 **"微信读书 - Word 2019专业增强风格"** 脚本处于开启状态。

4. **开始阅读**：
   - 打开或刷新 [微信读书网页版](https://weread.qq.com/) 页面，即刻进入 Word 伪装模式。

---**效果图**：
<img width="3840" height="1907" alt="image" src="https://github.com/user-attachments/assets/adba323d-cb2f-4a5f-a33d-03afb056dd27" />

## 🔄 更新日志

- **v1.0.0** - 当前版本：初始发布，包含核心拟真 UI、老板键与全局快捷键等完美摸鱼功能。

---

## ⚠️ 注意事项

- 如果你之前安装过此脚本的旧版本，**请务必在油猴面板中删除或禁用旧版本**。
- 如果更新后界面排版显示异常，请尝试使用 <kbd>Ctrl</kbd> + <kbd>F5</kbd> 强制刷新页面或清除浏览器缓存。
- 本脚本仅在本地修改前端界面的 CSS 与 DOM 结构，**绝对不收集、不上传任何用户数据**，请放心使用。

---

## 🤝 问题反馈与致谢

- 本项目基于 [odysseyw/WeReadFaker](https://github.com/odysseyw/WeReadFaker) 进行二次深度开发与重构，在此对原作者的开源精神表示感谢！
- 如在使用过程中遇到布局错位或 Bug，欢迎通过 GitHub Issue 提交反馈。

---

## 📄 许可证

本项目基于 **MIT 许可证** 开源，允许自由修改和二次发布，但请保留原作者及本项目的版权声明。
