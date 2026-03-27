# PPTX Merger

A lightweight, local desktop application built with Python and Tkinter that allows users to easily merge multiple Microsoft PowerPoint (`.pptx` / `.ppt`) presentations into a single file. 

Designed specifically for offline, closed-network environments, this tool processes files entirely locally without requiring internet access or cloud APIs.

## ✨ Features
* **Drag and Drop Interface:** Easily add files by dragging them directly into the application window.
* **Custom Reordering:** Drag items within the list or use the Up/Down buttons to set the exact order of the merged presentation.
* **Live Progress Tracking:** Visual progress bar and status updates during the merge process.
* **Completely Offline:** Uses local Windows COM automation. No data ever leaves the machine.
* **Hebrew UI:** Native RTL-friendly interface designed for Hebrew-speaking users.

## ⚠️ Prerequisites
Because this application uses COM automation to process the presentations natively, the target Windows machine **must** have the following installed:
* **Windows OS** (Windows 10 / 11)
* **Microsoft PowerPoint** (Installed and activated)
