# 🎓 Tutor Manager Pro+

![Python](https://img.shields.io/badge/python-3.10%2B-blue?style=for-the-badge&logo=python)
![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet?style=for-the-badge)
![SQLite](https://img.shields.io/badge/Database-SQLite3-lightgrey?style=for-the-badge&logo=sqlite)
![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)

> A sleek, powerful, and offline-first management desktop application built for private tutors, instructors, and small educational centers. Track your students, schedule lessons, and calculate your income effortlessly!

---

## 📑 Table of Contents
- [✨ Features](#-features)
- [🚀 Installation](#-installation)
- [💻 Usage](#-usage)
- [📂 Project Structure](#-project-structure)
- [🤝 Contributing](#-contributing)

---

## ✨ Features

### 🎨 Modern & Intuitive Interface
- **Dark Mode Native:** Beautifully designed using `CustomTkinter` to reduce eye strain during late-night scheduling.  
- **Cross-Platform:** Works seamlessly on Windows and macOS.

### 📅 Smart Scheduling
- **Interactive Calendar:** View your monthly and daily schedule at a glance.  
- **Dynamic Highlighting:** Days with scheduled lessons are automatically highlighted.  
- **Quick Add:** Instantly add one-off lessons or makeup classes.

### 💰 Financial & Data Tracking
- **Automated Income Calculation:** Set hourly rates per student/group and let the app calculate your earnings.  
- **Excel Reports:** Generate detailed, styled `.xlsx` reports grouped by month, group, and student with a single click.

### 🔒 Security & Reliability
- **Local Storage:** 100% offline. Your data stays securely on your machine using SQLite.  
- **1-Click Backup:** Easily create and restore database backups directly from the app.

---

## 🚀 Installation

Follow these steps to get the app running locally:

```bash
git clone https://github.com/HarKats03/Tutor_App.git
cd Tutor_App
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

---

## 💻 Usage

```bash
python tutor_app.py
```

💡 **First Run Note:**  
The app will automatically generate:
- `tutor_manager.db` (database)
- `logo.png`, `logo.ico` (icons)

---

## 📂 Project Structure

```plaintext
Tutor_App/
├── tutor_app.py             # Main application and GUI logic
├── requirements.txt         # Project dependencies
├── tutor_manager.db         # SQLite database (auto-created)
├── tutor_manager_backup.db  # Database backup (created via app)
└── README.md                # Project documentation
```

---

## 🤝 Contributing

1. Fork the repository  
2. Create your branch:
```bash
git checkout -b feature/AmazingFeature
```
3. Commit your changes:
```bash
git commit -m "Add some AmazingFeature"
```
4. Push to GitHub:
```bash
git push origin feature/AmazingFeature
```
5. Open a Pull Request 🚀

---

✨ Enjoy using **Tutor Manager Pro+**!