# Verizon Packing Tools

A suite of desktop tools for processing Verizon incoming BAU reel loads:
building packing lists from RITM emails, creating scrap pickup requests and
reel summary reports, splitting material weights, filling the BAU payment
statement, and appending scanned delivery slips.

**Version:** 1.0.0
**Platform:** Windows 10 / 11 (64-bit)
**Requirements:** Microsoft Excel and Microsoft Outlook (desktop versions)

---

## Table of Contents

1. [Purpose](#purpose)
2. [Getting Started](#getting-started)
3. [The Four Tools](#the-four-tools)
   - [Packing List](#1-packing-list)
   - [Split Material](#2-split-material)
   - [Fill Payment Statement](#3-fill-payment-statement)
   - [PDF Merge](#4-pdf-merge)
4. [What You Still Need To Do By Hand](#what-you-still-need-to-do-by-hand)
5. [The Database](#the-database)
6. [Quirks and Known Limitations](#quirks-and-known-limitations)
7. [Logs and Troubleshooting](#logs-and-troubleshooting)
8. [Folder Layout](#folder-layout)
9. [Frequently Asked Questions](#frequently-asked-questions)

---

## Purpose

This application simplifies organizing and making packing lists for **BAU
Loads** (specifically **REELS**, not loose scrap). It:

- Grabs the RITM from Outlook emails
- Grabs the attached PDF and sets up the per-load folder
- Creates a Scrap Pick Up Request per load
- Creates and fills the "Reel Summary Report WE"
- Inputs the data into a SQLite file
- Splits Material weights across a family of loads, based on RITM and
  location, and writes those into the summary workbook

---

## Getting Started

### First launch

1. Double-click **`VerizonPackingTools.exe`**.
2. The launcher opens with buttons for each tool.
3. The **first time** you click **Open Input Sheet** in Packing List, the
   blank template will land in a folder like:
