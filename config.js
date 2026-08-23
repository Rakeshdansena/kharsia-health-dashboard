// Kharsia Health Dashboard - Google Sheet Live CSV Links
const SHEET_BASE_URL = "https://docs.google.com/spreadsheets/d/1XAGjeCrLSVzTIraRSGkkjejXlrJEn-G2GxUEnN6ZCI0/export?format=csv&gid=";

const REPORT_CONFIG = {
  ayushman: {
    title: "Ayushman Card & Vay Vandana Report",
    csvUrl: SHEET_BASE_URL + "0"
  },
  rch: {
    title: "RCH 2.0 PW Registration Detail",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ RCH का gid बदलें
  },
  ncd: {
    title: "NCD Status 2026-27 Report",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ NCD का gid बदलें
  },
  jas: {
    title: "JAS Meeting Reporting",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ JAS Meeting का gid बदलें
  },
  hwc: {
    title: "Health & Wellness Center - Block Kharsia",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ HWC का gid बदलें
  },
  shivir: {
    title: "Ayushman Shivir Reporting FY 2026-27",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ Shivir Report का gid बदलें
  },
  activity: {
    title: "Ayushman Arogya Mandir Wellness Activity",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ Wellness Activity का gid बदलें
  },
  rbsk: {
    title: "RBSK Screening Report",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ RBSK का gid बदलें
  },
  nrc: {
    title: "NRC Kharsia Admission & Cure Report",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ NRC का gid बदलें
  },
  blindness: {
    title: "Blindness Control (NBCP) Progress",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ Blindness का gid बदलें
  },
  nqas: {
    title: "NQAS Certification Status",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ NQAS का gid बदलें
  },
  dialysis: {
    title: "PM Dialysis Progress Status",
    csvUrl: SHEET_BASE_URL + "1536656599" // Dialysis का gid पहले से सेट है
  },
  nlep: {
    title: "NLEP Leprosy Eradication Report",
    csvUrl: SHEET_BASE_URL + "0" // यहाँ NLEP का gid बदलें
  }
};
