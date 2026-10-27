# Bookstall Operations

A web-based operations and inventory management system built with Google Apps Script and HTML/JavaScript, designed to track book distribution, stock allocation, and volunteer sales across stalls and venues.

This Apps Script project is attached to the [Bookstall Operations Dev](https://docs.google.com/spreadsheets/d/14lzdc0ODb7EmTmRoHqa01kXg54xD7m0Cp02bdVxZiDY/edit?gid=0#gid=0).

---

## Overview

Bookstall Operations V1 streamlines the end-to-end workflow for running distributed bookstalls. It provides a centralized interface for volunteers and coordinators to review inventory, log sales, manage targets, and batch-allocate book stocks using automated matching.

---

## Key Features

* **Bulk Stock Allocation & Review**
  * Paste unstructured stock lists directly into the allocation tool.
  * Automatic title matching against the standard book catalog.
  * Discrepancy detection with manual title selection overrides and quantity adjustments.
* **Live Inventory Tracking**
  * Real-time visibility into book quantities by owner, location, and availability status.
  * Filter inventory by owner and stock status.
* **Activity & Distribution History**
  * Searchable historical records of all allocations, movements, and volunteer actions.
  * Multi-dimensional filtering by location, recorded volunteer/owner, and month.
* **Monthly Targets & Performance**
  * Month-by-month target setting and progress tracking for locations and stalls.
* **Google Workspace Integration**
  * Backed directly by Google Sheets for structured data persistence.
  * Native execution inside Google Apps Script with zero external hosting overhead.

---

## Architecture & File Structure

```text
├── Code.gs          # Server-side Apps Script: spreadsheet handlers, data commit APIs, and routing
├── Index.html       # Client-side UI: styles, layout, interactive tables, and event listeners
└── README.md        # Project documentation