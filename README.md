# AP Police TS Correspondence & Tappals Analytics Dashboard

> **Andhra Pradesh Police • Police Computer Services & Standardization (PCS&S)**  
> Technical Services Correspondence, Tappals & Task Tracking System

---

## 📊 Overview

An interactive, executive-grade Correspondence & Tappals Analytics Dashboard for Andhra Pradesh Police Technical Services. It tracks inbound communications across technical projects (CCTNS, NATGRID, AI4AP, Police Website, ICJS, Data Center, etc.), measuring officer workload, turnaround times (TAT), and task resolution stages.

### 🌟 Key Capabilities
- **Executive KPI Cards**: Real-time tracking of Total Inbound (48), Pending (22), In-Progress (14), Closed/Dispatched (12), Critical & High Alerts (8), and Avg Turnaround Time (12.3 Days).
- **Interactive Visualizations**: Officer workload distribution, project portfolios, communication channel mix, and correspondence aging radar.
- **Officers & Staff Hub**: Workload tracking and technical staff allocation for each officer.
- **Live Records Explorer**: Instant multi-column filtering (Officer, Project, Priority, Stage, Channel), search, and detail modal.
- **Dual Mode (Local & Web)**: 
  - **Local Mode**: Real-time synchronization with `AP TS-Correspondence MATRIX.xlsx` via `server.py` or `Launch_Dashboard.bat`.
  - **Static / Web Mode (GitHub Pages)**: Fully functional in-browser with drag-and-drop Excel file upload & pre-compiled analytics.

---

## 🚀 Quick Start (Local)

1. Double-click `Launch_Dashboard.bat` on Windows.
2. The dashboard will automatically open in your default browser at `http://localhost:8050`.
3. When updates are saved in `AP TS-Correspondence MATRIX.xlsx`, the dashboard updates automatically!

---

## 🌐 Live Web Access (GitHub Pages)

When hosted on GitHub Pages:
- The dashboard runs entirely client-side.
- You can click **"Upload Excel"** in the top bar to load or update any correspondence sheet directly in your browser.

---

## 📂 Project Structure

```
├── AP TS-Correspondence MATRIX.xlsx   # Master Correspondence Matrix Excel file
├── index.html                         # Executive Dashboard Single Page Application
├── server.py                          # Dynamic Python server with real-time Excel parsing
├── data.json                          # Pre-compiled static data snapshot for web deployment
├── Launch_Dashboard.bat               # 1-Click Windows launcher
└── README.md                          # Documentation
```
