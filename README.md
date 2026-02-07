# 🍃 Desktop Zen (净·桌面)

**The Ultimate Minimalist Desktop Cleaner for Windows**
<br>
*极致轻量 · 强制收纳 · 禅意归一*

![](https://img.shields.io/badge/Language-C%2B%2B17-00599C?style=flat-square&logo=c%2B%2B&logoColor=white)
![](https://img.shields.io/badge/Platform-Win7%2F10%2F11-0078D6?style=flat-square&logo=windows&logoColor=white)
![](https://img.shields.io/badge/Size-300KB-success?style=flat-square)

---

## 📖 Introduction

**DesktopZen** is a background utility written in pure C++ and Win32 API.
It enforces a "Zen Policy" on your desktop: **Keep it clean.**

Unlike other bloated desktop organizers, DesktopZen runs invisibly with near-zero resource usage.

**净·桌面** 是一款用纯 C++ 编写的 Windows 桌面净化工具。
它会监控你的桌面，当文件过多时，自动将它们移入收纳盒（缓存），只保留最新的文件。
还你一个没有任何杂念的屏幕。

## ✨ Features

*   🚀 **Ultra Lightweight:** Less than 200KB executable size.
*   ⚡ **Zero Latency:** Uses `ReadDirectoryChangesW` for kernel-level monitoring.
*   🕰️ **Legacy Support:** Compatible with Windows 7 (NT 6.1) kernel.
*   🛡️ **Zen Mode:** Auto-cleaning logic with a "Regret" (Undo) hotkey (`Alt+Z`).

## 🛠️ Build

Just open `DesktopZen.cpp` in Visual Studio and compile with `C++17`.
No external dependencies required.

---

<div align="center">
Made with ❤️ by Linkium[https://github.com/Linkium-suki]
</div>
