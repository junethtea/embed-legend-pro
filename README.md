# 🗺️ Embed Legend Pro for QGIS

[![QGIS Plugin](https://img.shields.io/badge/QGIS-Plugin-589632.svg)](https://plugins.qgis.org/plugins/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Donate](https://img.shields.io/badge/Donate-Buy_Me_a_Coffee-FF813F.svg)](#support--donate)

**Stop wasting time manually cropping legends and fixing messy Google Earth exports!**


**Embed Legend Pro** is the ultimate productivity booster designed specifically for RF Engineers, GIS Analysts, and Drive Test Coordinators. It creates a live, floating legend panel right on your QGIS map canvas and offers advanced thematic export capabilities (KMZ & MIF) that native tools lack.

Built by an RF Engineer, for RF Engineers. 📶

---

## 🚀 Key Features

* **Interactive Floating Panel:** A draggable legend that stays on top of your map canvas. Move it anywhere!
* **Smart Visibility Toggle:** Click any item in the legend to instantly hide/show features on the map with a "strikethrough" visual effect.
* **Multi-Layer Stacking:** Select multiple layers, and the legend automatically stacks them with clear folder-like separators.
* **Real-Time Statistics:** Automatically calculates feature counts and percentages (%) for every thematic category.
* **Professional KMZ Export (Google Earth):**
  * **Auto-Embedded Legend:** Exports the KMZ with the legend image overlay fixed perfectly on the screen.
  * **Clean Geometry:** Points (Drive Test) and Lines (Routes) are exported *clean* without annoying text labels cluttering the view.
  * **Smart Labeling:** Automatically detects `SiteID`, `eNB`, or `CellName` columns to generate professional "Yellow Floating Labels" *only* for Sector/Polygon layers.
* **Blazing Fast:** Export thousands of Gcell/Site data points to Google Earth in less than a minute.
* **MapInfo Export:** Direct export to `.MIF/.MID` format with hardcoded thematics.
* **Customizable UI:** Choose between "Modern Box" or "Clean Minimalist" styles. Adjust fonts and colors to match your reporting needs.

---

## 📸 Screenshots

<img width="548" height="474" alt="screenshot_qgis" src="https://github.com/user-attachments/assets/907f2885-9d55-426b-ace8-0cd142b8e6c0" />
<img width="742" height="504" alt="screenshot_ge" src="https://github.com/user-attachments/assets/59eec4f3-68a7-4796-9d11-1e832c7daeb2" />

---

## 🛠️ Installation

**Method 1: Via QGIS Official Repository (Recommended)**
1. Open QGIS.
2. Go to `Plugins` > `Manage and Install Plugins...`
3. Search for **Embed Legend Pro**.
4. Click `Install Plugin`.

**Method 2: Manual Installation (ZIP)**
1. Download the latest `.zip` release from this repository.
2. Open QGIS, go to `Plugins` > `Manage and Install Plugins...` > `Install from ZIP`.
3. Select the downloaded ZIP file and install.

---

## 💡 How to Use
1. Select one or multiple Vector Layers in your QGIS Layers Panel.
2. Click the **Embed Legend** icon in the toolbar.
3. The floating legend will appear. Right-click the panel to customize the style, colors, and toggle count/percentages.
4. Right-click the panel and select **Export KMZ** or **Export MIF** to generate your professional reports.

---

## ☕ Support & Donate

This tool is provided 100% free and open-source. If this plugin saves you hours of manual work, helps you hit your optimization targets, or just makes your RF reporting life easier, consider buying me a coffee! Your support keeps this project alive and continuously improving.

* 🌍 **International:** [Buy Me a Coffee](https://www.buymeacoffee.com/juneth) or [PayPal](https://paypal.me/junethtea)
* 🇮🇩 **Indonesia (Local):** OVO / GoPay at `081510027058`

*“May this tool be useful and become a continuous charity (amal jariah), especially for my beloved late parents. 🤲”*

---

## 👨‍💻 Author
**Jujun Junaedi** Telecommunications & NPO Engineer | Python & GIS Enthusiast  
Email: jujun.junaedi@outlook.com or jujun.telco@gmail.com
