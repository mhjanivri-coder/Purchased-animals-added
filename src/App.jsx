import React, { useEffect, useMemo, useRef, useState } from "react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";

const BREEDS = ["Murrah buffalo", "Nili-Ravi buffalo"];
const SEX_OPTIONS = ["Female", "Male"];
const STATUS_OPTIONS = ["Active (present in herd)", "Dead", "Culled"];
const CALVING_OUTCOMES = ["", "Normal calving", "Stillbirth", "Premature calving", "Abortion"];
const ENTRY_MODES = ["Manual", "Friday Records"];
const AI_RESULTS = ["Pending", "Negative", "Conceived"];
const FEMALE_TABS = ["photographs", "pedigree", "reproduction", "calving", "production", "health", "timeline", "history"];
const MALE_TABS = ["photographs", "pedigree", "disease testing", "dam record", "ai information", "progenies born", "performance of daughters", "health", "overall history sheet"];
const HEALTH_SUBTABS = [
  { id: "bodyWeight", label: "Body Weight" },
  { id: "deworming", label: "Deworming" },
  { id: "vaccination", label: "Vaccination" },
  { id: "treatment", label: "Treatment" },
];
const COLOSTRUM_DAYS = 5;
const STORAGE_KEY = "dairyscope_frontend_animals_v1";

const ENTRY_SOURCE_OPTIONS = ["Already part of herd", "Purchased"];

const emptyAnimal = {
  tagNo: "",
  breed: "Nili-Ravi buffalo",
  dob: "",
  category: "Female",
  identificationMark: "",
  status: "Active (present in herd)",
  entrySource: "Already part of herd",
  entryDate: "",
  exitDate: "",
  exitReason: "",
  isBreedingBull: "No",
  breedingSet: "",
  currentSet: "Yes",
  isPtBull: "No",
};

const emptyPedigree = {
  sire: "",
  dam: "",
  sireSire: "",
  sireDam: "",
  damSire: "",
  damDam: "",
  sireSireSire: "",
  sireSireDam: "",
  sireDamSire: "",
  sireDamDam: "",
  damSireSire: "",
  damSireDam: "",
  damDamSire: "",
  damDamDam: "",
};


function normalizeTagForLookup(value) {
  return String(value || "").trim().toLowerCase();
}

function getPedigreeObjectForAnimal(animal) {
  if (!animal) return emptyPedigree;
  if (animal.category === "Female") return { ...emptyPedigree, ...(animal.femaleDetails?.pedigree || {}) };
  if (animal.category === "Male") return { ...emptyPedigree, ...(animal.maleDetails?.pedigree || {}) };
  return emptyPedigree;
}

function getAnimalPedigreeLookupKeys(animal) {
  if (!animal) return [];
  const keys = [animal.tagNo];
  const tag = String(animal.tagNo || "").trim();
  const set = String(animal.breedingSet || "").trim();
  if (tag && set) {
    keys.push(`${tag}/${set}`);
    keys.push(`${tag} / ${set}`);
    keys.push(`${tag} ${set}`);
  }
  return keys.map(normalizeTagForLookup).filter(Boolean);
}

function findAnimalByTagNo(animals, tagNo) {
  const key = normalizeTagForLookup(tagNo);
  if (!key) return null;
  return (animals || []).find((a) => getAnimalPedigreeLookupKeys(a).includes(key)) || null;
}

function firstNonBlank(...values) {
  for (const value of values) {
    const text = String(value || "").trim();
    if (text) return text;
  }
  return "";
}

function derivePedigreeFromKnownRelatives(animal, allAnimals, visited = new Set()) {
  const base = getPedigreeObjectForAnimal(animal);
  const result = { ...emptyPedigree, ...base };
  if (!animal) return result;

  const selfKey = normalizeTagForLookup(animal.tagNo || animal.id || "");
  if (selfKey) visited.add(selfKey);

  const sireAnimal = findAnimalByTagNo(allAnimals, result.sire);
  const damAnimal = findAnimalByTagNo(allAnimals, result.dam);

  const sireKey = normalizeTagForLookup(sireAnimal?.tagNo || "");
  const damKey = normalizeTagForLookup(damAnimal?.tagNo || "");
  const sirePed = sireAnimal && !visited.has(sireKey)
    ? derivePedigreeFromKnownRelatives(sireAnimal, allAnimals, new Set(visited))
    : getPedigreeObjectForAnimal(sireAnimal);
  const damPed = damAnimal && !visited.has(damKey)
    ? derivePedigreeFromKnownRelatives(damAnimal, allAnimals, new Set(visited))
    : getPedigreeObjectForAnimal(damAnimal);

  result.sireSire = firstNonBlank(result.sireSire, sirePed.sire);
  result.sireDam = firstNonBlank(result.sireDam, sirePed.dam);
  result.damSire = firstNonBlank(result.damSire, damPed.sire);
  result.damDam = firstNonBlank(result.damDam, damPed.dam);

  result.sireSireSire = firstNonBlank(result.sireSireSire, sirePed.sireSire);
  result.sireSireDam = firstNonBlank(result.sireSireDam, sirePed.sireDam);
  result.sireDamSire = firstNonBlank(result.sireDamSire, sirePed.damSire);
  result.sireDamDam = firstNonBlank(result.sireDamDam, sirePed.damDam);
  result.damSireSire = firstNonBlank(result.damSireSire, damPed.sireSire);
  result.damSireDam = firstNonBlank(result.damSireDam, damPed.sireDam);
  result.damDamSire = firstNonBlank(result.damDamSire, damPed.damSire);
  result.damDamDam = firstNonBlank(result.damDamDam, damPed.damDam);

  return result;
}

function makeCalvingOutcomeEvent(outcomeNo = 1) {
  return {
    outcomeNo: String(outcomeNo),
    calvingDate: "",
    calfSex: "",
    calfTag: "",
    calfSire: "",
    calvingOutcome: "",
    remarks: "",
  };
}

function makeCalvingParity(parityNo) {
  return {
    parityNo: String(parityNo),
    calvingDate: "",
    calfSex: "",
    calfTag: "",
    calfSire: "",
    calvingOutcome: "",
    remarks: "",
    outcomeEvents: [makeCalvingOutcomeEvent(1)],
  };
}

function makeReproParity(parityNo) {
  return {
    parityNo: String(parityNo),
    conceptionDate: "",
    expectedCalvingDate: "",
    remarks: "",
    aiRecords: [],
  };
}

function makeFridayRecord(date = "") {
  return {
    date,
    morningMilk: "",
    eveningMilk: "",
    totalDailyYield: "",
    fatPct: "",
    snfPct: "",
    tsPct: "",
  };
}

function makeProductionLactation(parityNo) {
  return {
    parityNo: String(parityNo),
    entryMode: "Manual",
    calvingDate: "",
    dryDate: "",
    manualSummary: {
      totalLactationMilk: "",
      standardLactationMilk: "",
      peakYield: "",
    },
    fridayRecords: [],
  };
}

function makeBodyWeightRecord() {
  return { recordDate: "", bodyWeight: "" };
}

function makeDewormingRecord() {
  return { dewormingDate: "", anthelminticUsed: "" };
}

function makeVaccinationRecord() {
  return { vaccinationDate: "", vaccineUsed: "" };
}

function makeTreatmentRecord() {
  return { treatmentDate: "", diagnosis: "", treatmentGiven: "" };
}

function makeDiseaseTestRecord() {
  return { testDate: "", testName: "", result: "", remarks: "" };
}

const emptyHealth = {
  bodyWeightRecords: [makeBodyWeightRecord()],
  dewormingRecords: [makeDewormingRecord()],
  vaccinationRecords: [makeVaccinationRecord()],
  treatmentRecords: [makeTreatmentRecord()],
};

const emptyFemaleDetails = {
  pedigree: { ...emptyPedigree },
  photographs: [],
  calvingParities: [makeCalvingParity(1)],
  reproductionParities: [makeReproParity(0)],
  selectedReproParity: "0",
  productionLactations: [],
  selectedProductionParity: "",
  health: { ...emptyHealth },
  historyMeta: {
    reasonForCulling: "",
    bookValue: "",
  },
};

const emptyMaleDetails = {
  pedigree: { ...emptyPedigree },
  photographs: [],
  diseaseTests: [makeDiseaseTestRecord()],
  health: { ...emptyHealth },
  historyMeta: {
    remarks: "",
    bookValue: "",
  },
};

const initialAnimals = [];

function parseDisplayDate(value) {
  if (!value || typeof value !== "string") return null;
  const parts = value.trim().split("/");
  if (parts.length !== 3) return null;
  const day = Number(parts[0]);
  const month = Number(parts[1]);
  const year = Number(parts[2]);
  if (!day || !month || !year) return null;
  const dt = new Date(year, month - 1, day);
  if (dt.getFullYear() !== year || dt.getMonth() !== month - 1 || dt.getDate() !== day) return null;
  return dt;
}

function formatDateDisplay(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = String(date.getFullYear());
  return `${dd}/${mm}/${yyyy}`;
}

function normalizeDisplayDate(value) {
  const dt = parseDisplayDate(value);
  return dt ? formatDateDisplay(dt) : value;
}

function daysBetween(start, end) {
  const a = parseDisplayDate(start);
  const b = parseDisplayDate(end);
  if (!a || !b) return 0;
  return Math.max(0, Math.round((b.getTime() - a.getTime()) / 86400000));
}

function addDays(dateStr, days) {
  const dt = parseDisplayDate(dateStr);
  if (!dt) return "";
  const copy = new Date(dt);
  copy.setDate(copy.getDate() + days);
  return formatDateDisplay(copy);
}

function iconButtonClass(tone = "emerald", size = "lg") {
  const palette = tone === "amber"
    ? "border-amber-300 text-amber-900 hover:bg-amber-50"
    : "border-emerald-300 text-emerald-900 hover:bg-emerald-50";
  const scale = size === "sm" ? "h-11 w-11 text-xl" : "h-12 w-12 text-2xl";
  return `inline-flex items-center justify-center rounded-2xl border bg-white/95 font-bold leading-none shadow-sm transition ${palette} ${scale}`;
}

function isFridayDisplayDate(value) {
  const dt = parseDisplayDate(value);
  return Boolean(dt) && dt.getDay() === 5;
}

function getNextFridayDisplay(date = new Date()) {
  const dt = new Date(date);
  dt.setHours(0, 0, 0, 0);
  const offset = (5 - dt.getDay() + 7) % 7;
  dt.setDate(dt.getDate() + offset);
  return formatDateDisplay(dt);
}

function parseNumericValue(value) {
  if (value === null || value === undefined || value === "") return "";
  const normalized = String(value).replace(/,/g, "").trim();
  if (!normalized) return "";
  const num = Number(normalized);
  return Number.isFinite(num) ? String(num) : "";
}

function parseMilkEntryValue(value) {
  if (value === null || value === undefined) return "";
  const normalized = String(value).replace(/,/g, "").trim();
  if (!normalized) return "";
  if (!/^\d*(?:\.\d*)?$/.test(normalized)) return "";
  return normalized;
}

function formatMilkOneDecimal(value) {
  if (value === null || value === undefined || value === "") return "";
  const num = Number(value);
  return Number.isFinite(num) ? num.toFixed(1) : "";
}

function splitDelimitedLine(line, delimiter) {
  const out = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delimiter && !inQuotes) {
      out.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  out.push(current);
  return out.map((cell) => cell.trim());
}

function parseDelimitedText(text) {
  const lines = String(text || "").replace(/^\ufeff/, "").split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) return [];
  const headerLine = lines[0];
  const delimiter = headerLine.includes("	") ? "	" : headerLine.includes(";") && !headerLine.includes(",") ? ";" : ",";
  const headers = splitDelimitedLine(headerLine, delimiter).map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const values = splitDelimitedLine(line, delimiter);
    return headers.reduce((acc, header, idx) => {
      acc[header] = values[idx] ?? "";
      return acc;
    }, {});
  });
}

function parseSpreadsheetFile(file, fileContent) {
  const name = String(file?.name || "").toLowerCase();
  const isExcel = name.endsWith(".xlsx") || name.endsWith(".xls");
  if (isExcel) {
    const workbook = XLSX.read(fileContent, { type: "array" });
    const firstSheetName = workbook.SheetNames?.[0];
    if (!firstSheetName) return [];
    const sheet = workbook.Sheets[firstSheetName];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" });
  }
  return parseDelimitedText(String(fileContent || ""));
}

function getCurrentOrLatestLactation(animal) {
  const lactations = animal?.femaleDetails?.productionLactations || [];
  if (!lactations.length) return null;
  return [...lactations].sort((a, b) => Number(b.parityNo) - Number(a.parityNo))[0] || null;
}

function getLactationForDate(animal, dateStr) {
  const target = parseDisplayDate(dateStr);
  if (!target || animal?.category !== "Female") return null;
  const lactations = [...(animal?.femaleDetails?.productionLactations || [])].sort((a, b) => Number(a.parityNo) - Number(b.parityNo));
  let matched = null;
  lactations.forEach((l) => {
    const calving = parseDisplayDate(l.calvingDate || "");
    if (!calving || calving.getTime() > target.getTime()) return;
    const dry = parseDisplayDate(l.dryDate || "");
    if (dry && dry.getTime() < target.getTime()) return;
    matched = l;
  });
  return matched || null;
}

function isFemaleInMilkOnDate(animal, dateStr) {
  return getFemaleLifecycleOnDate(animal, dateStr) === "Milk";
}

function countsAsDryOnDate(animal, dateStr) {
  const lifecycle = getFemaleLifecycleOnDate(animal, dateStr);
  return lifecycle === "Dry" || lifecycle === "Colostrum";
}

function countsAsHeiferOnDate(animal, dateStr) {
  const lifecycle = getFemaleLifecycleOnDate(animal, dateStr);
  return lifecycle === "Heifer" || lifecycle === "Colostrum-Heifer";
}

function getFridayMilkEntryAnimals(animals, dateStr) {
  return animals
    .filter((a) => a.category === "Female" && !isArchivedAnimal(a) && isFemaleInMilkOnDate(a, dateStr))
    .slice()
    .sort(sortByTag);
}

function getDryFemaleAnimalsOnDate(animals, dateStr) {
  const target = parseDisplayDate(dateStr);
  if (!target) return [];
  return animals.filter((a) => a.category === "Female" && !isArchivedAnimal(a) && countsAsDryOnDate(a, dateStr));
}

function buildWAFridayRow(animals, dateStr) {
  const milkAnimals = getFridayMilkEntryAnimals(animals, dateStr);
  const dryAnimals = getDryFemaleAnimalsOnDate(animals, dateStr);
  const totalMilk = milkAnimals.reduce((sum, animal) => {
    const lactation = getLactationForDate(animal, dateStr);
    const record = lactation?.fridayRecords?.find((r) => normalizeDisplayDate(r.date || "") === normalizeDisplayDate(dateStr));
    return sum + Number(record?.totalDailyYield || 0);
  }, 0);
  const milkCount = milkAnimals.length;
  const dryCount = dryAnimals.length;
  const herdCount = milkCount + dryCount;
  return {
    date: normalizeDisplayDate(dateStr),
    milkAnimals: milkCount,
    dryAnimals: dryCount,
    totalMilk: Number(totalMilk.toFixed(2)),
    wa: milkCount ? Number((totalMilk / milkCount).toFixed(2)) : 0,
    ha: herdCount ? Number((totalMilk / herdCount).toFixed(2)) : 0,
  };
}

function getAllFridaySummaryRows(animals) {
  const dates = new Set();
  animals.filter((a) => a.category === "Female").forEach((animal) => {
    (animal?.femaleDetails?.productionLactations || []).forEach((l) => {
      (l.fridayRecords || []).forEach((r) => {
        if (isFridayDisplayDate(r.date || "")) dates.add(normalizeDisplayDate(r.date || ""));
      });
    });
  });
  return [...dates]
    .sort((a, b) => {
      const ad = parseDisplayDate(a); const bd = parseDisplayDate(b);
      if (!ad || !bd) return 0;
      return bd.getTime() - ad.getTime();
    })
    .map((dateStr) => buildWAFridayRow(animals, dateStr));
}

function expectedCalving(dateStr) {
  return dateStr ? addDays(dateStr, 310) : "";
}

function firstRecordableFriday(calvingDate) {
  const base = parseDisplayDate(calvingDate);
  if (!base) return "";
  for (let i = 0; i <= 14; i += 1) {
    const d = new Date(base);
    d.setDate(d.getDate() + i);
    const candidate = formatDateDisplay(d);
    const gap = daysBetween(calvingDate, candidate);
    if (d.getDay() === 5 && gap > 5) return candidate;
  }
  return "";
}

function sortByTag(a, b) {
  const an = Number(a.tagNo);
  const bn = Number(b.tagNo);
  const aNum = Number.isFinite(an) && !Number.isNaN(an);
  const bNum = Number.isFinite(bn) && !Number.isNaN(bn);
  if (aNum && bNum) return an - bn;
  return String(a.tagNo).localeCompare(String(b.tagNo), undefined, { numeric: true, sensitivity: "base" });
}

function normalizeRomanInput(value) {
  return (value || "").toUpperCase().replace(/[^IVXLCDM]/g, "");
}

function isArchivedAnimal(animal) {
  const archivedStatus = animal?.status === "Dead" || animal?.status === "Culled";
  return archivedStatus && Boolean((animal?.exitDate || "").trim()) && Boolean((animal?.exitReason || "").trim());
}

function normalizeAnimalFormData(form) {
  const next = { ...form };
  next.entrySource = next.entrySource || "Already part of herd";
  if (next.entrySource === "Purchased" && !next.entryDate) {
    next.entryDate = formatDateDisplay(new Date());
  }
  if (next.entrySource !== "Purchased") {
    next.entryDate = "";
  }
  if (next.status === "Active (present in herd)") {
    next.exitDate = "";
    next.exitReason = "";
  }
  if (next.category !== "Male") {
    next.isBreedingBull = "No";
    next.breedingSet = "";
    next.currentSet = "No";
    next.isPtBull = "No";
  } else {
    next.isBreedingBull = next.isBreedingBull || "No";
    next.breedingSet = next.isBreedingBull === "Yes" ? normalizeRomanInput(next.breedingSet || "") : "";
    next.currentSet = next.isBreedingBull === "Yes" ? (next.currentSet || "Yes") : "No";
    next.isPtBull = next.isBreedingBull === "Yes" && next.currentSet === "No" ? (next.isPtBull || "No") : "No";
  }
  return next;
}

function femaleHasAnyCalving(animal) {
  return Boolean(animal?.femaleDetails?.calvingParities?.some((p) => p.calvingDate));
}

function getLastNormalCalving(animal) {
  const valid = getAllCalvingOutcomeEvents(animal)
    .filter((event) => isCompleteProductiveCalvingEvent(event))
    .sort((a, b) => {
      const ad = parseDisplayDate(a.calvingDate);
      const bd = parseDisplayDate(b.calvingDate);
      if (!ad || !bd) return 0;
      return bd.getTime() - ad.getTime();
    });
  return valid[0] || null;
}

function getCurrentLactationDryDate(animal) {
  const lactations = animal?.femaleDetails?.productionLactations || [];
  if (!lactations.length) return "";
  const sorted = [...lactations].sort((a, b) => Number(a.parityNo) - Number(b.parityNo));
  return sorted[sorted.length - 1]?.dryDate || "";
}

function getFemaleLifecycleOnDate(animal, dateStr) {
  if (!animal || animal.category !== "Female") return animal?.category || "";
  const targetDate = normalizeDisplayDate(dateStr || formatDateDisplay(new Date()));
  const target = parseDisplayDate(targetDate);
  if (!target) return "";

  const productiveCalvings = getAllCalvingOutcomeEvents(animal)
    .filter((event) => isCompleteProductiveCalvingEvent(event))
    .filter((event) => {
      const d = parseDisplayDate(event.calvingDate || "");
      return d && d.getTime() <= target.getTime();
    })
    .sort((a, b) => {
      const ad = parseDisplayDate(a.calvingDate || "");
      const bd = parseDisplayDate(b.calvingDate || "");
      if (!ad || !bd) return 0;
      return bd.getTime() - ad.getTime();
    });

  if (!productiveCalvings.length) return "Heifer";

  const lastCalving = productiveCalvings[0];
  const latestParityNo = Number(lastCalving?.parityNo || 0);
  const daysSinceCalving = daysBetween(lastCalving.calvingDate, targetDate);
  if (daysSinceCalving < COLOSTRUM_DAYS) {
    return latestParityNo <= 1 ? "Colostrum-Heifer" : "Colostrum";
  }

  const productionLactations = animal?.femaleDetails?.productionLactations || [];
  const currentParityLactation = productionLactations.find((l) => Number(l.parityNo) === latestParityNo);
  const latestLactation = currentParityLactation || getCurrentOrLatestLactation(animal);
  if (latestLactation?.dryDate) {
    const dry = parseDisplayDate(latestLactation.dryDate || "");
    if (dry && dry.getTime() <= target.getTime()) return "Dry";
  }

  const lactation = getLactationForDate(animal, targetDate);
  if (lactation?.dryDate) {
    const dry = parseDisplayDate(lactation.dryDate || "");
    if (dry && dry.getTime() <= target.getTime()) return "Dry";
  }

  return "Milk";
}


function getFemaleLifecycle(animal) {
  return getFemaleLifecycleOnDate(animal, formatDateDisplay(new Date()));
}

function ordinalDay(day) {
  const suffix = day % 10 === 1 && day % 100 !== 11 ? "st" : day % 10 === 2 && day % 100 !== 12 ? "nd" : day % 10 === 3 && day % 100 !== 13 ? "rd" : "th";
  return `${day}${suffix}`;
}

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function getFinancialMonthOptions() {
  return ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];
}

function getFinancialMonthDate(fyStartYear, monthLabel) {
  const monthIndex = MONTH_NAMES.indexOf(monthLabel);
  if (monthIndex < 0) return new Date(Number(fyStartYear || new Date().getFullYear()), 3, 1);
  const year = monthIndex >= 3 ? Number(fyStartYear) : Number(fyStartYear) + 1;
  return new Date(year, monthIndex, 1);
}

function isAnimalPresentOnDate(animal, date) {
  const dob = parseDisplayDate(animal?.dob || "");
  if (dob && dob.getTime() > date.getTime()) return false;
  const exitDate = parseDisplayDate(animal?.exitDate || "");
  if (exitDate && exitDate.getTime() < date.getTime()) return false;
  return true;
}

function ageDaysOnDate(animal, date) {
  const dob = parseDisplayDate(animal?.dob || "");
  if (!dob) return null;
  return Math.max(0, Math.floor((date.getTime() - dob.getTime()) / 86400000));
}

function ageMonthsOnDate(animal, date) {
  const days = ageDaysOnDate(animal, date);
  return days === null ? null : days / 30.4375;
}

function buildHerdStrengthReport(animals, fyStartYear, monthLabel) {
  const asOfDate = getFinancialMonthDate(fyStartYear, monthLabel);
  const asOfDateDisplay = formatDateDisplay(asOfDate);
  const present = (animals || []).filter((a) => {
    if (!isAnimalPresentOnDate(a, asOfDate)) return false;
    if (!isArchivedAnimal(a)) return true;
    const exitDate = parseDisplayDate(a.exitDate || "");
    return Boolean(exitDate && exitDate.getTime() >= asOfDate.getTime());
  });
  const calfDefs = [
    ["0-30 days", (days) => days !== null && days <= 30],
    ["31-60 days", (days) => days !== null && days >= 31 && days <= 60],
    ["61-90 days", (days) => days !== null && days >= 61 && days <= 90],
    ["91-180 days", (days) => days !== null && days >= 91 && days <= 180],
    ["06 -12 Month", (days) => days !== null && days >= 181 && days <= 365],
  ];

  const makeRow = (label, male, female) => ({ label, male, female, total: male + female });

  const calfRows = calfDefs.map(([label, test]) => {
    const male = present.filter((a) => a.category === "Male" && test(ageDaysOnDate(a, asOfDate))).length;
    const female = present.filter((a) => a.category === "Female" && test(ageDaysOnDate(a, asOfDate))).length;
    return makeRow(label, male, female);
  });

  const femaleLifecycle = (a) => getFemaleLifecycleOnDate(a, asOfDateDisplay);
  const lactating = present.filter((a) => a.category === "Female" && (ageDaysOnDate(a, asOfDate) || 0) > 365 && femaleLifecycle(a) === "Milk").length;
  const dry = present.filter((a) => a.category === "Female" && (ageDaysOnDate(a, asOfDate) || 0) > 365 && femaleLifecycle(a) === "Dry").length;
  const heifer = present.filter((a) => a.category === "Female" && (ageMonthsOnDate(a, asOfDate) || 0) > 36 && femaleLifecycle(a) === "Heifer").length;
  const breedingBulls = present.filter((a) => a.category === "Male" && a.isBreedingBull === "Yes").length;
  const teaserBulls = present.filter((a) => a.category === "Male" && String(a.maleType || a.type || "").toLowerCase().includes("teaser")).length;

  const inMonths = (a, min, max) => {
    const m = ageMonthsOnDate(a, asOfDate);
    return m !== null && m >= min && m <= max;
  };
  const regularMale = (a) => a.category === "Male" && a.isBreedingBull !== "Yes" && !String(a.maleType || a.type || "").toLowerCase().includes("teaser");
  const youngFemale = (a) => a.category === "Female" && femaleLifecycle(a) === "Heifer";

  const adultRows = [
    makeRow("Lactating Buffaloes", 0, lactating),
    makeRow("Dry Buffaloes", 0, dry),
    makeRow("Heifers", 0, heifer),
    makeRow("Breeding Bulls", breedingBulls, 0),
    makeRow("Teaser Bull", teaserBulls, 0),
    makeRow("13 -24 Months", present.filter((a) => regularMale(a) && inMonths(a, 13, 24)).length, present.filter((a) => youngFemale(a) && inMonths(a, 13, 24)).length),
    makeRow("25 -30 Months", present.filter((a) => regularMale(a) && inMonths(a, 25, 30)).length, present.filter((a) => youngFemale(a) && inMonths(a, 25, 30)).length),
    makeRow(">30-36 Months", present.filter((a) => regularMale(a) && inMonths(a, 30.0001, 36)).length, present.filter((a) => youngFemale(a) && inMonths(a, 30.0001, 36)).length),
  ];

  const total = (rows, key) => rows.reduce((sum, row) => sum + Number(row[key] || 0), 0);
  return {
    asOfDate,
    monthLabel,
    calfRows,
    adultRows,
    calfTotals: makeRow("Total", total(calfRows, "male"), total(calfRows, "female")),
    adultTotals: makeRow("Total", total(adultRows, "male"), total(adultRows, "female")),
  };
}

function recalcFridayRecord(record) {
  const hasMilkEntry = record.morningMilk !== "" || record.eveningMilk !== "";
  const total = Number(record.morningMilk || 0) + Number(record.eveningMilk || 0);
  return {
    ...record,
    morningMilk: record.morningMilk !== "" ? formatMilkOneDecimal(record.morningMilk) : "",
    eveningMilk: record.eveningMilk !== "" ? formatMilkOneDecimal(record.eveningMilk) : "",
    totalDailyYield: hasMilkEntry ? total.toFixed(1) : record.totalDailyYield || "",
  };
}

function getNextFridayRecordDate(lactation) {
  const existing = lactation?.fridayRecords || [];
  if (!existing.length) return firstRecordableFriday(lactation?.calvingDate || "");
  const lastDate = existing[existing.length - 1]?.date || "";
  return lastDate ? addDays(lastDate, 7) : "";
}

function getHealthWithDefaults(health) {
  return {
    ...emptyHealth,
    ...(health || {}),
    bodyWeightRecords: health?.bodyWeightRecords?.length ? health.bodyWeightRecords.map((r) => ({ ...r })) : [makeBodyWeightRecord()],
    dewormingRecords: health?.dewormingRecords?.length ? health.dewormingRecords.map((r) => ({ ...r })) : [makeDewormingRecord()],
    vaccinationRecords: health?.vaccinationRecords?.length ? health.vaccinationRecords.map((r) => ({ ...r })) : [makeVaccinationRecord()],
    treatmentRecords: health?.treatmentRecords?.length ? health.treatmentRecords.map((r) => ({ ...r })) : [makeTreatmentRecord()],
  };
}

function withDefaults(animal) {
  const femaleDetails = {
    ...emptyFemaleDetails,
    ...(animal.femaleDetails || {}),
    pedigree: { ...emptyPedigree, ...(animal.femaleDetails?.pedigree || {}) },
    photographs: Array.isArray(animal.femaleDetails?.photographs) ? animal.femaleDetails.photographs.slice(0, 10).map((p) => ({ ...p })) : [],
    calvingParities: animal.femaleDetails?.calvingParities?.length
      ? animal.femaleDetails.calvingParities.map((p) => ({
          ...makeCalvingParity(p.parityNo ?? 1),
          ...p,
          outcomeEvents: getCalvingOutcomeEvents(p),
        }))
      : [makeCalvingParity(1)],
    reproductionParities: animal.femaleDetails?.reproductionParities?.length
      ? animal.femaleDetails.reproductionParities.map((p) => syncReproConceptionFields({
          ...makeReproParity(p.parityNo ?? 0),
          ...p,
          aiRecords: (p.aiRecords || []).map((r) => ({ ...r, aiBullNo: r.aiBullNo ?? r.aiBull ?? "", aiSetNo: r.aiSetNo ?? "" })),
        }))
      : [makeReproParity(0)],
    productionLactations: animal.femaleDetails?.productionLactations?.length
      ? animal.femaleDetails.productionLactations.map((l) => ({
          ...l,
          fridayRecords: (l.fridayRecords || []).map((r) => recalcFridayRecord({ ...makeFridayRecord(r.date || ""), ...r, totalDailyYield: r.totalDailyYield ?? r.totalMilk ?? "" })),
        }))
      : [],
    health: getHealthWithDefaults(animal.femaleDetails?.health),
    historyMeta: { ...emptyFemaleDetails.historyMeta, ...(animal.femaleDetails?.historyMeta || {}) },
  };

  const maleDetails = {
    ...emptyMaleDetails,
    ...(animal.maleDetails || {}),
    pedigree: { ...emptyPedigree, ...(animal.maleDetails?.pedigree || {}) },
    photographs: Array.isArray(animal.maleDetails?.photographs) ? animal.maleDetails.photographs.slice(0, 10).map((p) => ({ ...p })) : [],
    diseaseTests: animal.maleDetails?.diseaseTests?.length ? animal.maleDetails.diseaseTests.map((r) => ({ ...r })) : [makeDiseaseTestRecord()],
    health: getHealthWithDefaults(animal.maleDetails?.health),
    historyMeta: { ...emptyMaleDetails.historyMeta, ...(animal.maleDetails?.historyMeta || {}) },
  };

  const calvingCount = femaleDetails.calvingParities.length;
  const reproMap = new Map(femaleDetails.reproductionParities.map((p) => [String(p.parityNo), p]));
  if (!reproMap.has("0")) reproMap.set("0", makeReproParity(0));
  for (let i = 1; i <= calvingCount; i += 1) {
    if (!reproMap.has(String(i))) reproMap.set(String(i), makeReproParity(i));
  }
  const reproductionParities = [...reproMap.values()].sort((a, b) => Number(a.parityNo) - Number(b.parityNo));

  const productiveCalvingParities = femaleDetails.calvingParities
    .filter((cp) => getCalvingOutcomeEvents(cp).some((event) => isCompleteProductiveCalvingEvent(event)))
    .map((cp) => String(cp.parityNo));

  const lactationMap = new Map(femaleDetails.productionLactations.map((p) => [String(p.parityNo), p]));
  productiveCalvingParities.forEach((parityNo) => {
    if (!lactationMap.has(String(parityNo))) lactationMap.set(String(parityNo), makeProductionLactation(parityNo));
  });
  const productionLactations = [...lactationMap.values()]
    .filter((p) => productiveCalvingParities.includes(String(p.parityNo)))
    .sort((a, b) => Number(a.parityNo) - Number(b.parityNo))
    .map((l) => {
      const calvingDate = getCalvingOutcomeEvents(femaleDetails.calvingParities.find((c) => String(c.parityNo) === String(l.parityNo)))
        .find((event) => isCompleteProductiveCalvingEvent(event))?.calvingDate || "";
      return { ...l, calvingDate };
    });

  const selectedReproParity = reproductionParities.some((p) => String(p.parityNo) === String(femaleDetails.selectedReproParity)) ? String(femaleDetails.selectedReproParity) : "0";
  const selectedProductionParity = productionLactations.some((p) => String(p.parityNo) === String(femaleDetails.selectedProductionParity)) ? String(femaleDetails.selectedProductionParity) : productionLactations[productionLactations.length - 1]?.parityNo || "";

  return {
    ...animal,
    preCalvingLifecycle: animal.preCalvingLifecycle || "Heifer",
    exitDate: animal.exitDate || "",
    exitReason: animal.exitReason || "",
    isBreedingBull: animal.category === "Male" ? animal.isBreedingBull || "No" : "No",
    breedingSet: animal.category === "Male" && animal.isBreedingBull === "Yes" ? normalizeRomanInput(animal.breedingSet || "") : "",
    currentSet: animal.category === "Male" && animal.isBreedingBull === "Yes" ? (animal.currentSet || "Yes") : "No",
    isPtBull: animal.category === "Male" && animal.isBreedingBull === "Yes" && (animal.currentSet || "Yes") === "No" ? (animal.isPtBull || "No") : "No",
    femaleDetails: {
      ...femaleDetails,
      reproductionParities,
      selectedReproParity,
      productionLactations,
      selectedProductionParity,
    },
    maleDetails,
  };
}

function getSireStringFromAnimal(animal) {
  if (!animal) return "";
  return animal.category === "Female" ? animal.femaleDetails?.pedigree?.sire || "" : animal.maleDetails?.pedigree?.sire || "";
}

function isProgenyOfBull(progeny, bull) {
  const sire = getSireStringFromAnimal(progeny).trim();
  const bullTag = (bull?.tagNo || "").trim();
  if (!sire || !bullTag) return false;
  return sire === bullTag || sire.startsWith(`${bullTag}/`) || sire.startsWith(`${bullTag} `);
}

function buildAutoCalfAnimal(dam, calvingParity, outcomeEvent = null, outcomeIndex = 0) {
  if (dam?.category !== "Female") return null;
  const event = outcomeEvent || getCalvingOutcomeEvents(calvingParity).find((e) => isCalfIdentityOutcome(e));
  if (!isCalfIdentityOutcome(event)) return null;
  const calfTag = (event?.calfTag || "").trim();
  const calfSex = event?.calfSex || "";
  const calfDob = event?.calvingDate || "";
  const calfSire = (event?.calfSire || getCalfSireForCalving(dam, calvingParity?.parityNo, outcomeIndex) || "").trim();
  if (!calfTag || !calfSex || !calfDob) return null;

  const base = {
    id: `calf-${dam.id}-${calvingParity.parityNo}-${event.outcomeNo || outcomeIndex + 1}`,
    tagNo: calfTag,
    breed: dam.breed || "Nili-Ravi buffalo",
    dob: calfDob,
    category: calfSex === "Female" ? "Female" : "Male",
    identificationMark: "",
    status: "Active (present in herd)",
    exitDate: "",
    exitReason: "",
    isBreedingBull: "No",
    breedingSet: "",
    currentSet: "No",
    isPtBull: "No",
    linkedDamId: dam.id,
    linkedCalvingParityNo: String(calvingParity.parityNo),
    linkedCalvingOutcomeNo: String(event.outcomeNo || outcomeIndex + 1),
    autoAddedFromBirth: true,
  };

  if (calfSex === "Female") {
    return withDefaults({
      ...base,
      femaleDetails: {
        ...emptyFemaleDetails,
        pedigree: { ...emptyPedigree, dam: dam.tagNo || "", sire: calfSire },
      },
    });
  }

  return withDefaults({
    ...base,
    maleDetails: {
      ...emptyMaleDetails,
      pedigree: { ...emptyPedigree, dam: dam.tagNo || "", sire: calfSire },
    },
  });
}


function syncDamCalvesInHerd(animals, dam) {
  if (!dam || dam.category !== "Female") return animals;
  const calfRecords = (dam.femaleDetails?.calvingParities || []).flatMap((cp) =>
    getCalvingOutcomeEvents(cp).map((event, idx) => buildAutoCalfAnimal(dam, cp, event, idx)).filter(Boolean)
  );

  let nextAnimals = animals.filter((animal) => {
    if (!animal?.autoAddedFromBirth || animal?.linkedDamId !== dam.id) return true;
    return calfRecords.some((calf) => calf.id === animal.id);
  });

  calfRecords.forEach((calf) => {
    const existingIdx = nextAnimals.findIndex((animal) => animal.id === calf.id || (animal.tagNo === calf.tagNo && animal.id !== dam.id));
    if (existingIdx >= 0) {
      nextAnimals[existingIdx] = withDefaults({
        ...nextAnimals[existingIdx],
        ...calf,
        femaleDetails: calf.category === "Female" ? calf.femaleDetails : nextAnimals[existingIdx].femaleDetails,
        maleDetails: calf.category === "Male" ? calf.maleDetails : nextAnimals[existingIdx].maleDetails,
      });
    } else {
      nextAnimals = [calf, ...nextAnimals];
    }
  });

  return nextAnimals;
}

function nextDetailTab(tab, tabs) {
  const idx = tabs.indexOf(tab);
  return idx >= 0 && idx < tabs.length - 1 ? tabs[idx + 1] : tab;
}

function getSelectedReproParity(animal) {
  return animal?.femaleDetails?.reproductionParities?.find((p) => String(p.parityNo) === String(animal?.femaleDetails?.selectedReproParity)) || null;
}

function getSelectedLactation(animal) {
  return animal?.femaleDetails?.productionLactations?.find((p) => String(p.parityNo) === String(animal?.femaleDetails?.selectedProductionParity)) || null;
}

function getCalvingDateForParity(animal, parityNo) {
  return getProductiveCalvingEventForParity(animal, parityNo)?.calvingDate || "";
}

function isProductiveCalvingOutcome(row) {
  return ["Normal calving", "Stillbirth", "Premature calving"].includes(row?.calvingOutcome || "");
}

function isCalfIdentityOutcome(row) {
  return ["Normal calving", "Premature calving"].includes(row?.calvingOutcome || "");
}

function isNormalCalvingOutcome(row) {
  // Backward-compatible name: now means a calving event that should advance parity,
  // calculate AFC/CI, and open production/lactation parity.
  return isProductiveCalvingOutcome(row);
}

function isPregnancyClosingOutcome(row) {
  return ["Normal calving", "Stillbirth", "Premature calving", "Abortion"].includes(row?.calvingOutcome || "");
}

function getCalvingOutcomeEvents(calvingParity) {
  if (!calvingParity) return [];
  const rawEvents = Array.isArray(calvingParity.outcomeEvents) && calvingParity.outcomeEvents.length
    ? calvingParity.outcomeEvents
    : [{
        outcomeNo: "1",
        calvingDate: calvingParity.calvingDate || "",
        calfSex: calvingParity.calfSex || "",
        calfTag: calvingParity.calfTag || "",
        calfSire: calvingParity.calfSire || "",
        calvingOutcome: calvingParity.calvingOutcome || "",
        remarks: calvingParity.remarks || "",
      }];

  return rawEvents.map((event, idx) => ({
    ...makeCalvingOutcomeEvent(idx + 1),
    ...event,
    outcomeNo: String(event.outcomeNo || idx + 1),
  }));
}

function getAllCalvingOutcomeEvents(animal) {
  return (animal?.femaleDetails?.calvingParities || []).flatMap((cp) =>
    getCalvingOutcomeEvents(cp).map((event, idx) => ({
      ...event,
      parityNo: String(cp.parityNo),
      outcomeIndex: idx,
    }))
  );
}

function isReadyToFinalizeProductiveCalvingEvent(event) {
  if (!event?.calvingDate || !parseDisplayDate(event.calvingDate || "") || !isProductiveCalvingOutcome(event)) return false;
  if (isCalfIdentityOutcome(event)) {
    return Boolean(String(event.calfSex || "").trim() && String(event.calfTag || "").trim());
  }
  return true; // Stillbirth requires date + outcome only.
}

function isCompleteProductiveCalvingEvent(event) {
  return Boolean(event?.finalized && isReadyToFinalizeProductiveCalvingEvent(event));
}

function getCalvingRecordForParity(animal, parityNo) {
  const p = Number(parityNo);
  if (p <= 0) return null;
  return animal?.femaleDetails?.calvingParities?.find((c) => Number(c.parityNo) === p) || null;
}

function getProductiveCalvingEventsForParity(animal, parityNo) {
  const cp = getCalvingRecordForParity(animal, parityNo);
  return getCalvingOutcomeEvents(cp)
    .map((event, idx) => ({ ...event, parityNo: String(parityNo), outcomeIndex: idx }))
    .filter((event) => isCompleteProductiveCalvingEvent(event))
    .sort((a, b) => parseDisplayDate(a.calvingDate).getTime() - parseDisplayDate(b.calvingDate).getTime());
}

function getProductiveCalvingEventForParity(animal, parityNo) {
  const events = getProductiveCalvingEventsForParity(animal, parityNo);
  return events[0] || null;
}

function getNormalCalvingRecordForParity(animal, parityNo) {
  return getProductiveCalvingEventForParity(animal, parityNo);
}

function getNormalCalvingDateForParity(animal, parityNo) {
  return getProductiveCalvingEventForParity(animal, parityNo)?.calvingDate || "";
}

function getFirstNormalCalving(animal) {
  const rows = getAllCalvingOutcomeEvents(animal)
    .filter((row) => isCompleteProductiveCalvingEvent(row))
    .sort((a, b) => parseDisplayDate(a.calvingDate).getTime() - parseDisplayDate(b.calvingDate).getTime());
  return rows[0] || null;
}

function getCurrentCalvingParity(animal) {
  return (animal?.femaleDetails?.calvingParities || [])
    .filter((cp) => getProductiveCalvingEventForParity(animal, cp.parityNo)).length;
}

function getPregnancyOutcomeAfterDate(animal, dateStr) {
  const dt = parseDisplayDate(dateStr || "");
  if (!dt) return null;
  const rows = getAllCalvingOutcomeEvents(animal)
    .filter((row) => row.calvingDate && isPregnancyClosingOutcome(row))
    .filter((row) => {
      const outcomeDate = parseDisplayDate(row.calvingDate || "");
      return outcomeDate && outcomeDate.getTime() >= dt.getTime();
    })
    .sort((a, b) => parseDisplayDate(a.calvingDate).getTime() - parseDisplayDate(b.calvingDate).getTime());
  return rows[0] || null;
}

function pregnancyClosedByOutcome(animal, conceptionDate) {
  return Boolean(getPregnancyOutcomeAfterDate(animal, conceptionDate));
}

function getConceivedAIRecords(reproParity) {
  return sortAIRecords(reproParity?.aiRecords || []).filter((r) => r.result === "Conceived");
}

function getConceptionRecordForOutcome(animal, calvingParityNo, outcomeIndex = 0) {
  const reproParity = getReproParityByNo(animal, Number(calvingParityNo) - 1);
  const conceived = getConceivedAIRecords(reproParity);
  return conceived[outcomeIndex] || conceived[conceived.length - 1] || null;
}

function getConceptionDateForOutcome(animal, calvingParityNo, outcomeIndex = 0) {
  return getConceptionRecordForOutcome(animal, calvingParityNo, outcomeIndex)?.aiDate || "";
}

function getFirstAIForConceptionCycle(animal, calvingParityNo, outcomeIndex = 0) {
  const reproParity = getReproParityByNo(animal, Number(calvingParityNo) - 1);
  const aiRecords = sortAIRecords(reproParity?.aiRecords || []);
  const conceived = getConceptionRecordForOutcome(animal, calvingParityNo, outcomeIndex);
  if (!conceived?.aiDate) return aiRecords[0]?.aiDate || "";
  const previousOutcomeDate = outcomeIndex > 0
    ? getCalvingOutcomeEvents(getCalvingRecordForParity(animal, calvingParityNo))[outcomeIndex - 1]?.calvingDate
    : "";
  const start = parseDisplayDate(previousOutcomeDate || "");
  const end = parseDisplayDate(conceived.aiDate || "");
  const candidates = aiRecords.filter((r) => {
    const d = parseDisplayDate(r.aiDate || "");
    if (!d || !end) return false;
    return d.getTime() <= end.getTime() && (!start || d.getTime() > start.getTime());
  });
  return candidates[0]?.aiDate || conceived.aiDate;
}

function getAIServicesForConceptionCycle(animal, calvingParityNo, outcomeIndex = 0) {
  const reproParity = getReproParityByNo(animal, Number(calvingParityNo) - 1);
  const aiRecords = sortAIRecords(reproParity?.aiRecords || []);
  const conceived = getConceptionRecordForOutcome(animal, calvingParityNo, outcomeIndex);
  if (!conceived?.aiDate) return [];
  const previousOutcomeDate = outcomeIndex > 0
    ? getCalvingOutcomeEvents(getCalvingRecordForParity(animal, calvingParityNo))[outcomeIndex - 1]?.calvingDate
    : "";
  const start = parseDisplayDate(previousOutcomeDate || "");
  const end = parseDisplayDate(conceived.aiDate || "");
  return aiRecords.filter((r) => {
    const d = parseDisplayDate(r.aiDate || "");
    if (!d || !end) return false;
    return d.getTime() <= end.getTime() && (!start || d.getTime() > start.getTime());
  });
}

function getConceptionDateForReproParity(reproParity) {
  if (!reproParity) return "";
  const conceivedAI = getConceivedAIRecord(reproParity);
  return reproParity.conceptionDate || conceivedAI?.aiDate || "";
}

function hasOpenPregnancyInParity(animal, reproParity) {
  const conceptionDate = getConceptionDateForReproParity(reproParity);
  if (!conceptionDate) return false;
  return !pregnancyClosedByOutcome(animal, conceptionDate);
}

function getActiveExpectedCalving(animal, reproParity) {
  const conceptionDate = getConceptionDateForReproParity(reproParity);
  if (!conceptionDate) return "";
  if (pregnancyClosedByOutcome(animal, conceptionDate)) return "";
  return reproParity?.expectedCalvingDate || expectedCalving(conceptionDate);
}

function getProductiveCalvingRecordForReproParity(animal, reproParityNo) {
  const calvingParityNo = Number(reproParityNo) + 1;
  return getProductiveCalvingEventForParity(animal, calvingParityNo);
}

function isReproParityClosedByProductiveCalving(animal, reproParityNo) {
  return Boolean(getProductiveCalvingRecordForReproParity(animal, reproParityNo));
}

function syncProductionLactationsWithCalvingOutcomes(animal) {
  if (!animal || animal.category !== "Female") return animal;
  const femaleDetails = animal.femaleDetails || {};
  const calvingParities = femaleDetails.calvingParities || [];
  const existingLactations = femaleDetails.productionLactations || [];

  const productiveCalvings = calvingParities
    .map((cp) => {
      const event = getCalvingOutcomeEvents(cp).find((e) => isCompleteProductiveCalvingEvent(e));
      return event ? { ...event, parityNo: String(cp.parityNo) } : null;
    })
    .filter(Boolean)
    .sort((a, b) => Number(a.parityNo || 0) - Number(b.parityNo || 0));

  const productiveParityNos = productiveCalvings.map((cp) => String(cp.parityNo));

  let nextLactations = existingLactations
    .filter((lactation) => productiveParityNos.includes(String(lactation.parityNo)))
    .map((lactation) => {
      const cp = productiveCalvings.find((row) => String(row.parityNo) === String(lactation.parityNo));
      return { ...lactation, calvingDate: cp?.calvingDate || lactation.calvingDate || "" };
    });

  productiveCalvings.forEach((cp) => {
    const exists = nextLactations.some((l) => String(l.parityNo) === String(cp.parityNo));
    if (!exists) {
      nextLactations.push({ ...makeProductionLactation(cp.parityNo), calvingDate: cp.calvingDate || "" });
    }
  });

  nextLactations = nextLactations.sort((a, b) => Number(a.parityNo || 0) - Number(b.parityNo || 0));

  const selectedProductionParity = nextLactations.some((l) => String(l.parityNo) === String(femaleDetails.selectedProductionParity))
    ? femaleDetails.selectedProductionParity
    : (nextLactations[nextLactations.length - 1]?.parityNo || "");

  return withDefaults({
    ...animal,
    femaleDetails: {
      ...femaleDetails,
      productionLactations: nextLactations,
      selectedProductionParity,
    },
  });
}


function getReproParityByNo(animal, parityNo) {
  return animal?.femaleDetails?.reproductionParities?.find((p) => Number(p.parityNo) === Number(parityNo)) || null;
}

function hasMeaningfulReproContent(parity) {
  return Boolean(
    parity?.conceptionDate ||
    parity?.expectedCalvingDate ||
    parity?.remarks ||
    (parity?.aiRecords || []).some((r) => r.aiDate || r.aiBullNo || r.aiSetNo || (r.result && r.result !== "Pending"))
  );
}

function shouldShowReproParity(animal, parity) {
  const no = Number(parity?.parityNo || 0);
  if (no === 0) return true;
  if (hasMeaningfulReproContent(parity)) return true;
  if (isReproParityClosedByProductiveCalving(animal, no)) return true;
  // Repro parity n should appear only after productive calving parity n has closed.
  return Boolean(getProductiveCalvingEventForParity(animal, no));
}


function sortAIRecords(aiRecords = []) {
  return [...aiRecords].filter((r) => r.aiDate).sort((a, b) => {
    const ad = parseDisplayDate(a.aiDate);
    const bd = parseDisplayDate(b.aiDate);
    if (!ad || !bd) return 0;
    return ad.getTime() - bd.getTime();
  });
}

function getConceivedAIRecord(reproParity) {
  if (!reproParity) return null;
  const aiRecords = sortAIRecords(reproParity.aiRecords || []);
  const conceived = aiRecords.filter((r) => r.result === "Conceived");
  return conceived[conceived.length - 1] || null;
}

function syncReproConceptionFields(reproParity) {
  const conceivedAI = getConceivedAIRecord(reproParity);
  const conceptionDate = conceivedAI?.aiDate || "";
  return {
    ...reproParity,
    conceptionDate,
    expectedCalvingDate: conceptionDate ? expectedCalving(conceptionDate) : "",
  };
}

function formatBullSet(aiRecord) {
  if (!aiRecord) return "";
  const bullNo = (aiRecord.aiBullNo || "").trim();
  const setNo = (aiRecord.aiSetNo || "").trim();
  if (bullNo && setNo) return `${bullNo}/${setNo}`;
  return bullNo || setNo || "";
}

function getCalfSireForCalving(animal, calvingParityNo, outcomeIndex = 0) {
  const sourceReproParity = Number(calvingParityNo) - 1;
  if (sourceReproParity < 0) return "";
  const reproParity = getReproParityByNo(animal, sourceReproParity);
  const conceivedRecords = getConceivedAIRecords(reproParity);
  return formatBullSet(conceivedRecords[outcomeIndex] || conceivedRecords[conceivedRecords.length - 1] || getConceivedAIRecord(reproParity));
}

function cleanPedValue(v) {
  return String(v || "").trim().toUpperCase();
}

function canonicalPedId(value) {
  const raw = cleanPedValue(value);
  if (!raw) return "";
  // Treat bull/set notation such as 702/V as the same animal as tag 702 for relationship checks.
  // The full text is still displayed in pedigree fields; only matching is normalized here.
  return raw.split("/")[0].trim();
}

function getAnimalSelfIds(animal) {
  const ids = new Set();
  const tag = canonicalPedId(animal?.tagNo);
  if (tag) ids.add(tag);
  const setNo = cleanPedValue(animal?.breedingSet || "");
  if (tag && setNo) ids.add(canonicalPedId(`${tag}/${setNo}`));
  return ids;
}

function getPedigreeMapForMate(animal, allAnimals = []) {
  const p = derivePedigreeFromKnownRelatives(animal, allAnimals || []);
  return {
    self: canonicalPedId(animal?.tagNo),
    sire: canonicalPedId(p.sire),
    dam: canonicalPedId(p.dam),
    sireSire: canonicalPedId(p.sireSire),
    sireDam: canonicalPedId(p.sireDam),
    damSire: canonicalPedId(p.damSire),
    damDam: canonicalPedId(p.damDam),
    sireSireSire: canonicalPedId(p.sireSireSire),
    sireSireDam: canonicalPedId(p.sireSireDam),
    sireDamSire: canonicalPedId(p.sireDamSire),
    sireDamDam: canonicalPedId(p.sireDamDam),
    damSireSire: canonicalPedId(p.damSireSire),
    damSireDam: canonicalPedId(p.damSireDam),
    damDamSire: canonicalPedId(p.damDamSire),
    damDamDam: canonicalPedId(p.damDamDam),
  };
}

function buildAncestorDepthMap(ped) {
  const fields = [
    ["sire", 1, "sire"], ["dam", 1, "dam"],
    ["sireSire", 2, "sire's sire"], ["sireDam", 2, "sire's dam"],
    ["damSire", 2, "dam's sire"], ["damDam", 2, "dam's dam"],
    ["sireSireSire", 3, "sire-sire-sire"], ["sireSireDam", 3, "sire-sire-dam"],
    ["sireDamSire", 3, "sire-dam-sire"], ["sireDamDam", 3, "sire-dam-dam"],
    ["damSireSire", 3, "dam-sire-sire"], ["damSireDam", 3, "dam-sire-dam"],
    ["damDamSire", 3, "dam-dam-sire"], ["damDamDam", 3, "dam-dam-dam"],
  ];
  const map = new Map();
  fields.forEach(([field, depth, label]) => {
    const id = ped[field];
    if (!id) return;
    if (!map.has(id)) map.set(id, { id, depths: [], labels: [] });
    map.get(id).depths.push(depth);
    map.get(id).labels.push(label);
  });
  return map;
}

function lowestDepth(entry) {
  return Math.min(...(entry?.depths || [99]));
}

function estimateRelationshipScore(female, bull, allAnimals = []) {
  const f = getPedigreeMapForMate(female, allAnimals);
  const b = getPedigreeMapForMate(bull, allAnimals);

  let inbreeding = 0;
  const reasons = [];

  if (!bull?.tagNo || bull?.category !== "Male" || bull?.isBreedingBull !== "Yes") {
    return { score: 999, inbreedingPct: "High", status: "Avoid", reasons: ["Not marked as breeding bull"] };
  }

  const femaleSelfIds = getAnimalSelfIds(female);
  const bullSelfIds = getAnimalSelfIds(bull);
  const femaleAncestors = buildAncestorDepthMap(f);
  const bullAncestors = buildAncestorDepthMap(b);

  // Absolute blocks: parent-offspring, animal with itself, or direct ancestor/descendant.
  if ([...bullSelfIds].some((id) => femaleSelfIds.has(id))) {
    return { score: 999, inbreedingPct: "100.0", status: "Avoid", reasons: ["Same animal cannot be selected for mating"] };
  }
  if ([...bullSelfIds].some((id) => id && (id === f.sire || id === f.dam))) {
    return { score: 999, inbreedingPct: "25.0", status: "Avoid", reasons: ["Bull is direct parent of selected female"] };
  }
  if ([...femaleSelfIds].some((id) => id && (id === b.sire || id === b.dam))) {
    return { score: 999, inbreedingPct: "25.0", status: "Avoid", reasons: ["Bull is direct son/offspring of selected female"] };
  }
  if ([...bullSelfIds].some((id) => femaleAncestors.has(id))) {
    return { score: 999, inbreedingPct: "High", status: "Avoid", reasons: ["Bull is present in the female's pedigree"] };
  }
  if ([...femaleSelfIds].some((id) => bullAncestors.has(id))) {
    return { score: 999, inbreedingPct: "High", status: "Avoid", reasons: ["Selected female is present in the bull's pedigree"] };
  }

  // Wright's expected inbreeding of the proposed calf: Σ(1/2)^(n1+n2+1), assuming unknown ancestors are non-inbred.
  femaleAncestors.forEach((fEntry, id) => {
    const bEntry = bullAncestors.get(id);
    if (!bEntry) return;
    const n1 = lowestDepth(fEntry);
    const n2 = lowestDepth(bEntry);
    const contribution = Math.pow(0.5, n1 + n2 + 1) * 100;
    inbreeding += contribution;
    const relationName = n1 <= 1 && n2 <= 1 ? "parent" : n1 <= 2 && n2 <= 2 ? "grandparent" : "ancestor";
    reasons.push(`Common ${relationName} ${id} contributes ${contribution.toFixed(3)}%`);
  });

  let status = "Preferred";
  if (inbreeding >= 6.25) status = "Avoid";
  else if (inbreeding > 0) status = "Use with caution";

  return {
    score: inbreeding,
    inbreedingPct: inbreeding.toFixed(3),
    status,
    reasons: reasons.length ? reasons : ["No common ancestor detected within available 3-generation pedigree"],
  };
}

function getPreviouslyUsedBullKeys(female) {
  if (!female || female.category !== "Female") return new Set();
  const keys = new Set();
  (female.femaleDetails?.reproductionParities || []).forEach((parity) => {
    (parity.aiRecords || []).forEach((r) => {
      const bullNo = cleanPedValue(r.aiBullNo || "");
      const setNo = cleanPedValue(r.aiSetNo || "");
      if (bullNo || setNo) keys.add(`${bullNo}__${setNo}`);
    });
  });
  return keys;
}

function makeBullKey(bull) {
  return `${cleanPedValue(bull?.tagNo)}__${cleanPedValue(bull?.breedingSet)}`;
}

function getBestStandardLactationMilk(female) {
  if (!female || female.category !== "Female") return 0;
  return Math.max(
    0,
    ...((female.femaleDetails?.productionLactations || []).map((lactation) => {
      if (!lactation) return 0;
      if (lactation.entryMode === "Manual") {
        return Number(lactation.manualSummary?.standardLactationMilk || 0) || 0;
      }
      return Number(computeProductionMetrics(lactation).standardLactationMilk || 0) || 0;
    }))
  );
}

function isEliteFemaleForPtBull(female) {
  return getBestStandardLactationMilk(female) > 3000;
}

function isBullEligibleForFemaleMateSelection(bull, female) {
  if (!bull || bull.category !== "Male" || bull.isBreedingBull !== "Yes" || isArchivedAnimal(bull)) return false;
  if ((bull.currentSet || "Yes") === "Yes") return true;
  if ((bull.currentSet || "Yes") === "No" && bull.isPtBull === "Yes" && isEliteFemaleForPtBull(female)) return true;
  return false;
}

function classifyMateSelectionAdvanced(female, bull, allAnimals = []) {
  const relationship = estimateRelationshipScore(female, bull, allAnimals);
  const usedKeys = getPreviouslyUsedBullKeys(female);
  const currentBullKey = makeBullKey(bull);
  const reasons = [...relationship.reasons];
  if ((bull.currentSet || "Yes") === "Yes") {
    reasons.unshift("Current set bull");
  } else if (bull.isPtBull === "Yes") {
    reasons.unshift(`PT bull for elite female; best SLMY ${getBestStandardLactationMilk(female).toFixed(1)} kg`);
  }
  let status = relationship.status;
  let sortScore = relationship.score;

  if (usedKeys.has(currentBullKey)) {
    reasons.push("Already used in previous parity/AI records");
    sortScore += 100;
    if (status === "Preferred") status = "Use with caution";
    if (relationship.score >= 12.5) status = "Avoid";
  }

  return {
    bull,
    status,
    sortScore,
    inbreedingPct: relationship.inbreedingPct,
    reason: reasons.join("; "),
    alreadyUsed: usedKeys.has(currentBullKey),
  };
}

function getMateSelectionRowsAdvanced(female, animals) {
  if (!female || female.category !== "Female") return [];
  const rows = animals
    .filter((a) => isBullEligibleForFemaleMateSelection(a, female))
    .map((bull) => classifyMateSelectionAdvanced(female, bull, animals))
    .sort((a, b) => {
      if (a.sortScore !== b.sortScore) return a.sortScore - b.sortScore;
      return String(a.bull.tagNo).localeCompare(String(b.bull.tagNo), undefined, { numeric: true, sensitivity: "base" });
    });
  let usableRank = 0;
  return rows.map((row) => {
    if (row.status === "Avoid") return { ...row, autoRank: "Blocked", recommendation: "Do not use" };
    usableRank += 1;
    return {
      ...row,
      autoRank: usableRank,
      recommendation: usableRank === 1 ? "Best available" : usableRank <= 3 ? "Next option" : "Lower priority",
    };
  });
}


function getYearFromDisplayDate(value) {
  const dt = parseDisplayDate(value);
  return dt ? dt.getFullYear() : null;
}

function getMonthIndexFromDisplayDate(value) {
  const dt = parseDisplayDate(value);
  return dt ? dt.getMonth() : null;
}

function normKey(value) {
  return String(value || "").trim().toUpperCase();
}

function getFemaleGroupForConception(parityNo) {
  return Number(parityNo) === 0 ? "Heifer" : "Milk+ dry";
}

function getAllAIEvents(animals) {
  const rows = [];
  (animals || []).forEach((animal) => {
    if (animal?.category !== "Female") return;
    const femaleTag = animal.tagNo || "";
    const reproParities = animal?.femaleDetails?.reproductionParities || [];

    reproParities.forEach((parity) => {
      const parityNo = Number(parity?.parityNo || 0);
      const group = getFemaleGroupForConception(parityNo);

      const aiRecords = (parity?.aiRecords || [])
        .map((rec) => ({
          aiDate: rec?.aiDate || "",
          bullNo: String(rec?.aiBullNo || rec?.bullNo || "").trim(),
          setNo: String(rec?.aiSetNo || rec?.setNo || "").trim(),
          result: String(rec?.result || "Pending").trim(),
        }))
        .filter((rec) => rec.aiDate)
        .sort((a, b) => {
          const ad = parseDisplayDate(a.aiDate);
          const bd = parseDisplayDate(b.aiDate);
          if (!ad || !bd) return 0;
          return ad.getTime() - bd.getTime();
        });

      if (!aiRecords.length) return;

      // Backward-compatible fallback: if an older record only stored parity-level conceptionDate
      // and did not mark any AI row as Conceived, infer the most likely AI row.
      let fallbackConceivedIdx = -1;
      if (!aiRecords.some((rec) => rec.result === "Conceived") && parity?.conceptionDate) {
        const conceptionDt = parseDisplayDate(parity.conceptionDate || "");
        fallbackConceivedIdx = aiRecords.findIndex((rec) => rec.aiDate === parity.conceptionDate);
        if (fallbackConceivedIdx < 0 && conceptionDt) {
          aiRecords.forEach((rec, idx) => {
            const dt = parseDisplayDate(rec.aiDate);
            if (dt && dt.getTime() <= conceptionDt.getTime()) fallbackConceivedIdx = idx;
          });
        }
      }

      const bullAttemptCounter = new Map();
      aiRecords.forEach((rec, idx) => {
        const bullKey = `${normKey(rec.bullNo)}__${normKey(rec.setNo)}`;
        const bullAiNo = (bullAttemptCounter.get(bullKey) || 0) + 1;
        bullAttemptCounter.set(bullKey, bullAiNo);

        const isConceived = rec.result === "Conceived" || idx === fallbackConceivedIdx;

        rows.push({
          femaleId: animal.id,
          femaleTag,
          parityNo,
          group,
          aiNo: idx + 1,
          bullAiNo,
          aiDate: rec.aiDate,
          bullNo: rec.bullNo,
          setNo: rec.setNo,
          result: isConceived ? "Conceived" : rec.result,
          conceptionDate: isConceived ? (idx === fallbackConceivedIdx ? (parity?.conceptionDate || rec.aiDate) : rec.aiDate) : "",
        });
      });
    });
  });

  return rows.sort((a, b) => {
    const ad = parseDisplayDate(a.aiDate);
    const bd = parseDisplayDate(b.aiDate);
    if (!ad || !bd) return 0;
    return ad.getTime() - bd.getTime();
  });
}

function matchesBullRecord(aiRow, bullAnimal) {
  if (!aiRow || !bullAnimal) return false;
  const rowBull = normKey(aiRow?.bullNo);
  const bullTag = normKey(bullAnimal?.tagNo);
  if (!rowBull || !bullTag) return false;

  // Match by Bull No. only.
  // Historical AI records may contain the set used at the time of insemination
  // (e.g. 702 / Set V), while the bull registry may now show another/current set
  // (e.g. 702 / Set VII, PT). If we also force set matching, valid inseminations,
  // conceptions and progenies disappear from the bull AI Information tab.
  return rowBull === bullTag;
}

function getBullAIStats(aiRows) {
  const totalAI = aiRows.length;
  const conceived = aiRows.filter((r) => r.result === "Conceived").length;
  const negative = aiRows.filter((r) => r.result === "Negative").length;
  const pending = aiRows.filter((r) => r.result === "Pending").length;
  const crPct = totalAI ? ((conceived / totalAI) * 100).toFixed(2) : "0.00";

  const femaleParityMap = new Map();
  aiRows.forEach((row) => {
    const key = `${normKey(row?.femaleTag)}__${normKey(row?.parityNo)}`;
    if (!femaleParityMap.has(key)) femaleParityMap.set(key, []);
    femaleParityMap.get(key).push(row);
  });

  let femalesServed = 0;
  let firstAttemptConceived = 0;

  femaleParityMap.forEach((rows) => {
    const sortedRows = [...rows].sort((a, b) => {
      const ad = parseDisplayDate(a?.aiDate);
      const bd = parseDisplayDate(b?.aiDate);
      if (!ad || !bd) return 0;
      return ad.getTime() - bd.getTime();
    });
    if (!sortedRows.length) return;
    femalesServed += 1;
    if (sortedRows[0]?.result === "Conceived") firstAttemptConceived += 1;
  });

  const firstServiceConceptionRate = femalesServed
    ? ((firstAttemptConceived / femalesServed) * 100).toFixed(2)
    : "0.00";

  return {
    totalAI,
    conceived,
    negative,
    pending,
    crPct,
    femalesServed,
    firstAttemptConceived,
    firstServiceConceptionRate,
  };
}

function monthLabelFromIndex(idx) {
  return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][idx] || "";
}

function getFinancialYearOptions(animals) {
  const years = new Set();
  (animals || []).forEach((a) => {
    const dates = [a?.dob, a?.exitDate];
    (a?.femaleDetails?.calvingParities || []).forEach((c) => dates.push(c.calvingDate));
    (a?.femaleDetails?.reproductionParities || []).forEach((p) => {
      dates.push(p.conceptionDate);
      (p.aiRecords || []).forEach((r) => dates.push(r.aiDate));
    });
    (a?.femaleDetails?.productionLactations || []).forEach((l) => {
      dates.push(l.calvingDate);
      dates.push(l.dryDate);
      (l.fridayRecords || []).forEach((r) => dates.push(r.date));
    });
    dates.forEach((d) => {
      const yr = getYearFromDisplayDate(d);
      if (yr) {
        const fyStart = (parseDisplayDate(d).getMonth() >= 3) ? yr : yr - 1;
        years.add(fyStart);
      }
    });
  });
  if (!years.size) {
    const now = new Date();
    const fyStart = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
    years.add(fyStart);
  }
  return [...years].sort((a,b)=>b-a).map((start)=>({ value:String(start), label:`${start}-${String(start+1).slice(-2)}` }));
}

function getCalendarYearOptions(animals) {
  const years = new Set();
  getAllAIEvents(animals).forEach((r) => {
    const yr = getYearFromDisplayDate(r.aiDate);
    if (yr) years.add(yr);
  });
  if (!years.size) years.add(new Date().getFullYear());
  return [...years].sort((a,b)=>b-a).map((y)=>({ value:String(y), label:String(y) }));
}

function buildMonthlyFinancialSummary(animals, fyStartYear) {
  const rows = [];
  for (let offset = 0; offset < 12; offset += 1) {
    const monthIndex = (3 + offset) % 12;
    const year = monthIndex >= 3 ? fyStartYear : fyStartYear + 1;
    const row = { month: monthLabelFromIndex(monthIndex), monthIndex, year, calvings: 0, stillbirths: 0, abortions: 0, mortality: 0, lactationsCompleted: 0 };

    (animals || []).forEach((a) => {
      (a?.femaleDetails?.calvingParities || []).forEach((c) => {
        getCalvingOutcomeEvents(c).forEach((event) => {
          const dt = parseDisplayDate(event.calvingDate || "");
          if (!dt || dt.getFullYear() !== year || dt.getMonth() !== monthIndex) return;
          const outcome = event.calvingOutcome || "";
          if (outcome === "Normal calving" || outcome === "Premature calving") row.calvings += 1;
          if (outcome === "Stillbirth") row.stillbirths += 1;
          if (outcome === "Abortion") row.abortions += 1;
        });
      });

      if (isArchivedAnimal(a)) {
        const edt = parseDisplayDate(a.exitDate || "");
        if (edt && edt.getFullYear() === year && edt.getMonth() === monthIndex) row.mortality += 1;
      }

      (a?.femaleDetails?.productionLactations || []).forEach((l) => {
        const dt = parseDisplayDate(l.dryDate || "");
        if (dt && dt.getFullYear() === year && dt.getMonth() === monthIndex) row.lactationsCompleted += 1;
      });
    });

    rows.push(row);
  }
  return rows;
}

function getFinancialMonthYear(fyStartYear, monthLabel) {
  const monthMap = { Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11, Jan: 0, Feb: 1, Mar: 2 };
  const monthIndex = monthMap[monthLabel];
  if (monthIndex === undefined) return null;
  const year = monthIndex >= 3 ? fyStartYear : fyStartYear + 1;
  return { monthIndex, year };
}

function ageInMonthsBetween(startStr, endStr) {
  const start = parseDisplayDate(startStr);
  const end = parseDisplayDate(endStr);
  if (!start || !end) return "";
  const months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) - (end.getDate() < start.getDate() ? 1 : 0);
  return String(Math.max(0, months));
}

function buildMonthlyEventDetails(animals, fyStartYear, monthLabel, type) {
  const info = getFinancialMonthYear(fyStartYear, monthLabel);
  if (!info) return [];
  const { monthIndex, year } = info;
  const out = [];

  (animals || []).forEach((animal) => {
    if (animal?.category !== "Female" && type !== "mortality") return;

    if (type === "calvings") {
      (animal?.femaleDetails?.calvingParities || []).forEach((c) => {
        getCalvingOutcomeEvents(c).forEach((event, outcomeIdx) => {
          const dt = parseDisplayDate(event.calvingDate || "");
          if (!dt || dt.getFullYear() !== year || dt.getMonth() !== monthIndex) return;
          if (!["Normal calving", "Premature calving"].includes(event.calvingOutcome || "")) return;
          const metrics = computeCalvingOutcomeMetrics(animal, c.parityNo, event, outcomeIdx);
          out.push({
            tagNo: animal.tagNo || "",
            parityNo: String(c.parityNo || ""),
            calvingDate: event.calvingDate || "",
            calfSex: event.calfSex || "",
            calfTag: event.calfTag || "",
            calfSire: event.calfSire || getCalfSireForCalving(animal, c.parityNo, outcomeIdx) || "",
            afc: Number(c.parityNo) === 1 ? (metrics.afc ?? "") : "",
            sp: Number(c.parityNo) >= 2 ? (metrics.servicePeriod ?? "") : "",
            ci: Number(c.parityNo) >= 2 ? (metrics.calvingInterval ?? "") : "",
          });
        });
      });
    }

    if (type === "stillbirths") {
      (animal?.femaleDetails?.calvingParities || []).forEach((c) => {
        getCalvingOutcomeEvents(c).forEach((event, outcomeIdx) => {
          const dt = parseDisplayDate(event.calvingDate || "");
          if (!dt || dt.getFullYear() !== year || dt.getMonth() !== monthIndex) return;
          if (event.calvingOutcome !== "Stillbirth") return;
          out.push({
            tagNo: animal.tagNo || "",
            stillbirthDate: event.calvingDate || "",
            calfSire: event.calfSire || getCalfSireForCalving(animal, c.parityNo, outcomeIdx) || "",
          });
        });
      });
    }

    if (type === "abortions") {
      (animal?.femaleDetails?.calvingParities || []).forEach((c) => {
        getCalvingOutcomeEvents(c).forEach((event, outcomeIdx) => {
          const dt = parseDisplayDate(event.calvingDate || "");
          if (!dt || dt.getFullYear() !== year || dt.getMonth() !== monthIndex) return;
          if (event.calvingOutcome !== "Abortion") return;
          out.push({
            tagNo: animal.tagNo || "",
            conceptionDate: getConceptionDateForOutcome(animal, c.parityNo, outcomeIdx) || "",
            abortionDate: event.calvingDate || "",
            calfSire: event.calfSire || getCalfSireForCalving(animal, c.parityNo, outcomeIdx) || "",
          });
        });
      });
    }

    if (type === "lactationsCompleted") {
      (animal?.femaleDetails?.productionLactations || []).forEach((l) => {
        const dt = parseDisplayDate(l.dryDate || "");
        if (!dt || dt.getFullYear() !== year || dt.getMonth() !== monthIndex) return;
        const metrics = computeProductionMetrics(l);
        out.push({
          tagNo: animal.tagNo || "",
          parityNo: String(l.parityNo || ""),
          dryDate: l.dryDate || "",
          lactationLength: metrics.lactationLength || "",
          tlmy: l.entryMode === "Manual" ? (l.manualSummary.totalLactationMilk || "") : Number(metrics.totalLactationMilk || 0).toFixed(1),
          slmy: l.entryMode === "Manual" ? (l.manualSummary.standardLactationMilk || "") : Number(metrics.standardLactationMilk || 0).toFixed(1),
          peakYield: l.entryMode === "Manual" ? (l.manualSummary.peakYield || "") : Number(metrics.peakYield || 0).toFixed(1),
        });
      });
    }
  });

  if (type === "mortality") {
    (animals || []).forEach((animal) => {
      if (!isArchivedAnimal(animal)) return;
      const dt = parseDisplayDate(animal.exitDate || "");
      if (!dt || dt.getFullYear() !== year || dt.getMonth() !== monthIndex) return;
      out.push({
        tagNo: animal.tagNo || "",
        dob: animal.dob || "",
        deathDate: animal.exitDate || "",
        ageAtDeathMonths: ageInMonthsBetween(animal.dob || "", animal.exitDate || ""),
        cause: animal.exitReason || "",
      });
    });
  }

  return out.sort((a, b) => String(a.tagNo).localeCompare(String(b.tagNo), undefined, { numeric: true, sensitivity: "base" }));
}

function SummaryDetailTable({ detailType, rows, monthLabel, financialYearLabel, onClose }) {
  const titleMap = {
    calvings: "Calvings",
    stillbirths: "Stillbirths",
    abortions: "Abortions",
    mortality: "Mortality",
    lactationsCompleted: "Lactations Completed",
  };
  const title = titleMap[detailType] || "Details";

  return (
    <div className="mt-5 rounded-2xl border border-emerald-100 bg-white p-4">
      <div className="mb-2 flex items-center justify-between">
        <div className="text-xl font-semibold text-emerald-900">{title} · {monthLabel} · FY {financialYearLabel}</div>
        <button className="rounded-xl border border-emerald-300 bg-white px-3 py-2 text-sm text-emerald-800 hover:bg-emerald-50" onClick={onClose}>Close</button>
      </div>
      <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 p-1">
        {detailType === "calvings" && (
          <table className="min-w-full border-collapse text-sm">
            <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag Number","Parity","Date of calving","Sex of calf","Tag Number of calf","Sire of calf","AFC (days)","Service period","Calving interval"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
            <tbody>{rows.map((r, i)=><tr key={i}><td className="border border-emerald-100 px-2 py-2">{r.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{r.parityNo}</td><td className="border border-emerald-100 px-2 py-2">{r.calvingDate}</td><td className="border border-emerald-100 px-2 py-2">{r.calfSex}</td><td className="border border-emerald-100 px-2 py-2">{r.calfTag}</td><td className="border border-emerald-100 px-2 py-2">{r.calfSire}</td><td className="border border-emerald-100 px-2 py-2">{r.afc}</td><td className="border border-emerald-100 px-2 py-2">{r.sp}</td><td className="border border-emerald-100 px-2 py-2">{r.ci}</td></tr>)}</tbody>
          </table>
        )}
        {detailType === "stillbirths" && (
          <table className="min-w-full border-collapse text-sm">
            <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag Number","Date of stillbirth","Sire of calf"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
            <tbody>{rows.map((r, i)=><tr key={i}><td className="border border-emerald-100 px-2 py-2">{r.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{r.stillbirthDate}</td><td className="border border-emerald-100 px-2 py-2">{r.calfSire}</td></tr>)}</tbody>
          </table>
        )}
        {detailType === "abortions" && (
          <table className="min-w-full border-collapse text-sm">
            <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag Number","Date of conception","Date of abortion","Sire of calf"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
            <tbody>{rows.map((r, i)=><tr key={i}><td className="border border-emerald-100 px-2 py-2">{r.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{r.conceptionDate}</td><td className="border border-emerald-100 px-2 py-2">{r.abortionDate}</td><td className="border border-emerald-100 px-2 py-2">{r.calfSire}</td></tr>)}</tbody>
          </table>
        )}
        {detailType === "mortality" && (
          <table className="min-w-full border-collapse text-sm">
            <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag Number","Date of Birth","Date of death","Age at which died (months)","Cause of death"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
            <tbody>{rows.map((r, i)=><tr key={i}><td className="border border-emerald-100 px-2 py-2">{r.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{r.dob}</td><td className="border border-emerald-100 px-2 py-2">{r.deathDate}</td><td className="border border-emerald-100 px-2 py-2">{r.ageAtDeathMonths}</td><td className="border border-emerald-100 px-2 py-2">{r.cause}</td></tr>)}</tbody>
          </table>
        )}
        {detailType === "lactationsCompleted" && (
          <table className="min-w-full border-collapse text-sm">
            <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag Number","Current Parity","Date of Drying","Lactation length (days)","Total Lactation milk yield (kg)","Standard lactation Milk yield (kg)","Peak Yield (Kg)"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
            <tbody>{rows.map((r, i)=><tr key={i}><td className="border border-emerald-100 px-2 py-2">{r.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{r.parityNo}</td><td className="border border-emerald-100 px-2 py-2">{r.dryDate}</td><td className="border border-emerald-100 px-2 py-2">{r.lactationLength}</td><td className="border border-emerald-100 px-2 py-2">{r.tlmy}</td><td className="border border-emerald-100 px-2 py-2">{r.slmy}</td><td className="border border-emerald-100 px-2 py-2">{r.peakYield}</td></tr>)}</tbody>
          </table>
        )}
      </div>
      {rows.length === 0 && <div className="mt-3 text-sm text-slate-500">No animal records found for this month and section.</div>}
    </div>
  );
}

function buildMonthlyConceptionSummary(aiRows, calYear) {
  return Array.from({ length: 12 }, (_, idx) => {
    const totalAI = (aiRows || []).filter((r) => getYearFromDisplayDate(r.aiDate) === calYear && getMonthIndexFromDisplayDate(r.aiDate) === idx).length;
    const conceived = (aiRows || []).filter((r) => r.result === "Conceived" && getYearFromDisplayDate(r.conceptionDate || r.aiDate) === calYear && getMonthIndexFromDisplayDate(r.conceptionDate || r.aiDate) === idx).length;
    return { month: monthLabelFromIndex(idx), totalAI, conceived, crPct: totalAI ? ((conceived / totalAI) * 100).toFixed(2) : "0.00" };
  });
}

function buildMonthlyConceptionDetailRows(animals, aiRows, calYear, monthLabel) {
  const monthIdx = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(monthLabel);
  if (monthIdx < 0) return [];

  const animalMap = new Map((animals || []).map((animal) => [animal.id, animal]));

  const rowsByFemaleParity = new Map();
  (aiRows || []).forEach((row) => {
    const femaleId = row?.femaleId;
    if (!femaleId) return;
    const key = `${femaleId}__${Number(row?.parityNo || 0)}`;
    if (!rowsByFemaleParity.has(key)) rowsByFemaleParity.set(key, []);
    rowsByFemaleParity.get(key).push(row);
  });

  rowsByFemaleParity.forEach((rows) => {
    rows.sort((a, b) => {
      const ad = parseDisplayDate(a.aiDate);
      const bd = parseDisplayDate(b.aiDate);
      if (ad && bd && ad.getTime() !== bd.getTime()) return ad.getTime() - bd.getTime();
      return Number(a.aiNo || 0) - Number(b.aiNo || 0);
    });
  });

  return (aiRows || [])
    .filter((row) => getYearFromDisplayDate(row.aiDate) === calYear && getMonthIndexFromDisplayDate(row.aiDate) === monthIdx)
    .map((row) => {
      const animal = animalMap.get(row.femaleId);
      const parityNo = Number(row.parityNo || 0);
      const parityKey = `${row.femaleId}__${parityNo}`;
      const currentParityRows = rowsByFemaleParity.get(parityKey) || [];
      const parityAIIndex = currentParityRows.findIndex((r) => r.aiDate === row.aiDate && Number(r.aiNo || 0) === Number(row.aiNo || 0));
      const parityCalvingDate = parityNo > 0 ? getCalvingDateForParity(animal, parityNo) : "";
      return {
        tagNo: row.femaleTag || animal?.tagNo || "",
        lastCalvingDate: parityCalvingDate || "—",
        totalAISoFar: parityAIIndex >= 0 ? parityAIIndex + 1 : row.aiNo || "",
        bullUsed: [row.bullNo || "", row.setNo ? `Set ${row.setNo}` : ""].filter(Boolean).join(" / ") || "—",
        currentAIDate: row.aiDate || "",
        inseminationResult: row.result === "Conceived" ? "Conceived" : "Not-conceived",
      };
    })
    .sort((a, b) => {
      const ad = parseDisplayDate(a.currentAIDate);
      const bd = parseDisplayDate(b.currentAIDate);
      if (!ad || !bd) return String(a.tagNo).localeCompare(String(b.tagNo), undefined, { numeric: true, sensitivity: "base" });
      if (ad.getTime() !== bd.getTime()) return ad.getTime() - bd.getTime();
      return String(a.tagNo).localeCompare(String(b.tagNo), undefined, { numeric: true, sensitivity: "base" });
    });
}

function buildQuarterWiseConceptionSummary(aiRows, calYear) {
  const quarters = [
    { label: "Jan-Mar", months: [0, 1, 2] },
    { label: "Apr-Jun", months: [3, 4, 5] },
    { label: "Jul-Sep", months: [6, 7, 8] },
    { label: "Oct-Dec", months: [9, 10, 11] },
  ];
  return quarters.map((q) => {
    const totalAI = (aiRows || []).filter((r) => getYearFromDisplayDate(r.aiDate) === calYear && q.months.includes(getMonthIndexFromDisplayDate(r.aiDate))).length;
    const conceived = (aiRows || []).filter((r) => r.result === "Conceived" && getYearFromDisplayDate(r.conceptionDate || r.aiDate) === calYear && q.months.includes(getMonthIndexFromDisplayDate(r.conceptionDate || r.aiDate))).length;
    return { quarter: q.label, totalAI, conceived, crPct: totalAI ? ((conceived / totalAI) * 100).toFixed(2) : "0.00" };
  });
}

function buildBullWiseConceptionSummary(aiRows, calYear) {
  const map = new Map();
  (aiRows || []).forEach((row) => {
    if (getYearFromDisplayDate(row.aiDate) !== calYear) return;
    const key = `${normKey(row.bullNo)}__${normKey(row.setNo)}`;
    if (!map.has(key)) map.set(key, { bullNo: row.bullNo || "", setNo: row.setNo || "", totalAI: 0, conceived: 0 });
    const rec = map.get(key);
    rec.totalAI += 1;
    if (row.result === "Conceived" && getYearFromDisplayDate(row.conceptionDate || row.aiDate) === calYear) rec.conceived += 1;
  });
  return [...map.values()]
    .map((r) => ({ ...r, crPct: r.totalAI ? ((r.conceived / r.totalAI) * 100).toFixed(2) : "0.00" }))
    .sort((a, b) => String(a.bullNo).localeCompare(String(b.bullNo), undefined, { numeric: true, sensitivity: "base" }));
}

function buildFemaleConceptionByNthAI(aiRows, calYear) {
  const base = () => ({ ai: 0, conceived: 0, crPct: "0.00" });
  const out = {
    Heifer: { 1: base(), 2: base(), 3: base(), 4: base() },
    "Milk+ dry": { 1: base(), 2: base(), 3: base(), 4: base() },
  };
  (aiRows || []).forEach((r) => {
    if (getYearFromDisplayDate(r.aiDate) !== calYear) return;
    const grp = r.group === "Heifer" ? "Heifer" : "Milk+ dry";
    const nth = r.aiNo >= 4 ? 4 : r.aiNo;
    out[grp][nth].ai += 1;
    if (r.result === "Conceived" && getYearFromDisplayDate(r.conceptionDate || r.aiDate) === calYear) out[grp][nth].conceived += 1;
  });
  Object.values(out).forEach((groupObj) => {
    Object.values(groupObj).forEach((slot) => {
      slot.crPct = slot.ai ? ((slot.conceived / slot.ai) * 100).toFixed(2) : "0.00";
    });
  });
  return out;
}

function computeCalvingOutcomeMetrics(animal, calvingParityNo, outcomeEvent, outcomeIndex = 0) {
  const p = Number(calvingParityNo);
  const outcomeDate = outcomeEvent?.calvingDate || "";
  const previousProductiveCalving = getNormalCalvingDateForParity(animal, p - 1);
  const conceptionDate = getConceptionDateForOutcome(animal, p, outcomeIndex);

  let afc = null;
  const firstProductiveCalving = getFirstNormalCalving(animal);
  if (p === 1 && isCompleteProductiveCalvingEvent(outcomeEvent) && animal?.dob && firstProductiveCalving?.calvingDate && Number(firstProductiveCalving.parityNo || 0) === 1 && Number(firstProductiveCalving.outcomeIndex || 0) === Number(outcomeIndex)) {
    afc = daysBetween(animal.dob, firstProductiveCalving.calvingDate);
  }

  let gestationPeriod = null;
  if (conceptionDate && outcomeDate) {
    gestationPeriod = daysBetween(conceptionDate, outcomeDate);
  }

  let servicePeriod = null;
  // Official SP uses conception that resulted in productive calving only.
  if (p >= 2 && previousProductiveCalving && conceptionDate && isCompleteProductiveCalvingEvent(outcomeEvent)) {
    servicePeriod = daysBetween(previousProductiveCalving, conceptionDate);
  }

  let calvingInterval = null;
  if (p >= 2 && previousProductiveCalving && outcomeDate && isCompleteProductiveCalvingEvent(outcomeEvent)) {
    calvingInterval = daysBetween(previousProductiveCalving, outcomeDate);
  }

  return { afc, gestationPeriod, servicePeriod, calvingInterval };
}

function computeCalvingMetrics(animal, calvingParityNo) {
  const productive = getProductiveCalvingEventForParity(animal, calvingParityNo);
  const cp = getCalvingRecordForParity(animal, calvingParityNo);
  const firstEvent = productive || getCalvingOutcomeEvents(cp)[0] || makeCalvingOutcomeEvent(1);
  return computeCalvingOutcomeMetrics(animal, calvingParityNo, firstEvent, Number(firstEvent.outcomeIndex || 0));
}

function computeReproSummary(animal, parity) {
  if (!animal || !parity) return { parityLabel: "Heifer stage", lastAIDate: "", services: 0 };
  const parityNo = Number(parity.parityNo);
  const aiRecords = sortAIRecords(parity.aiRecords || []);
  const lastAIDate = aiRecords.length ? aiRecords[aiRecords.length - 1].aiDate : "";
  return { parityLabel: parityNo === 0 ? "Heifer stage" : `Parity ${parityNo}`, lastAIDate, services: aiRecords.length };
}

function getLatestCalvingParityNo(animal) {
  const calvings = getAllCalvingOutcomeEvents(animal)
    .filter((event) => isCompleteProductiveCalvingEvent(event))
    .sort((a, b) => Number(b.parityNo || 0) - Number(a.parityNo || 0));
  return calvings.length ? Number(calvings[0].parityNo || 0) : 0;
}

function hasCalvingAfterDate(animal, dateStr) {
  // Kept for compatibility: true when a pregnancy-closing outcome has occurred
  // after conception, including normal calving, stillbirth, or abortion.
  return pregnancyClosedByOutcome(animal, dateStr);
}

function getAnimalSummarySnapshot(animal) {
  if (!animal) return null;
  const ageDays = animal.dob ? daysBetween(animal.dob, formatDateDisplay(new Date())) : 0;
  const ageText = animal.dob && ageDays
    ? `${Math.floor(ageDays / 365)}y ${Math.floor((ageDays % 365) / 30)}m`
    : "—";

  if (animal.category !== "Female") {
    return {
      tagNo: animal.tagNo || "—",
      age: ageText,
      currentParity: "—",
      lastCalving: "—",
      lastAI: "—",
      pregnancyStatus: animal.isBreedingBull === "Yes" ? "Breeding bull" : "Male",
      expectedCalving: "—",
      milkStatus: "—",
      alerts: animal.status !== "Active (present in herd)" ? [`${animal.status}${animal.exitDate ? ` on ${animal.exitDate}` : ""}`] : [],
    };
  }

  const latestCalvingParityNo = getLatestCalvingParityNo(animal);
  const currentReproParityNo = latestCalvingParityNo || 0;
  const parities = animal.femaleDetails?.reproductionParities || [];
  const currentParity = parities.find((p) => Number(p.parityNo) === Number(currentReproParityNo)) || null;
  const aiRecords = sortAIRecords(currentParity?.aiRecords || []);
  const lastAI = aiRecords[aiRecords.length - 1] || null;
  const lastCalving = getLastNormalCalving(animal);
  const conceptionDate = getConceptionDateForReproParity(currentParity);
  const conceptionClosed = conceptionDate ? pregnancyClosedByOutcome(animal, conceptionDate) : false;
  const expected = getActiveExpectedCalving(animal, currentParity);
  const lifecycle = getFemaleLifecycle(animal);
  const alerts = [];

  if (lastAI?.aiDate) {
    const daysAfterAI = daysBetween(lastAI.aiDate, formatDateDisplay(new Date()));
    if (daysAfterAI > 30 && !["Conceived", "Negative"].includes(lastAI.result || "")) alerts.push("PD due");
  }
  if (expected) {
    const daysToCalving = daysBetween(formatDateDisplay(new Date()), expected);
    if (daysToCalving <= 30) alerts.push(daysToCalving < 0 ? "Calving overdue" : "Calving due");
  }
  if (aiRecords.length >= 4 && !conceptionDate) alerts.push("Repeat breeder");
  if (lastCalving?.calvingDate && aiRecords.length === 0 && daysBetween(lastCalving.calvingDate, formatDateDisplay(new Date())) > 90) alerts.push(">90d post-calving, no AI");
  if (lastCalving?.calvingDate && !conceptionDate && daysBetween(lastCalving.calvingDate, formatDateDisplay(new Date())) > 120) alerts.push(">120d post-calving, not pregnant");

  return {
    tagNo: animal.tagNo || "—",
    age: ageText,
    currentParity: latestCalvingParityNo > 0 ? `Parity ${latestCalvingParityNo}` : "Heifer stage",
    lastCalving: lastCalving?.calvingDate || "—",
    lastAI: lastAI?.aiDate || "—",
    pregnancyStatus: conceptionDate && !conceptionClosed ? "Pregnant / Conceived" : (lastAI ? (lastAI.result === "Negative" ? "Not pregnant" : conceptionClosed ? "Pregnancy closed" : "Pending PD") : "Not inseminated"),
    expectedCalving: expected || "—",
    milkStatus: lifecycle || "—",
    alerts,
  };
}

function buildAnimalTimeline(animal) {
  if (!animal) return [];
  const events = [];
  const add = (date, title, detail, tone = "slate") => {
    if (!date) return;
    events.push({ date, title, detail, tone });
  };

  add(animal.dob, "DOB", "Date of birth", "emerald");

  if (animal.category === "Female") {
    (animal.femaleDetails?.reproductionParities || []).forEach((parity) => {
      const aiRecords = sortAIRecords(parity.aiRecords || []);
      const conceivedRecord = aiRecords.find((rec) => rec.result === "Conceived");
      const conceptionDate = parity.conceptionDate || conceivedRecord?.aiDate || "";

      aiRecords.forEach((rec, idx) => {
        add(rec.aiDate, "AI", `Repro parity ${parity.parityNo} · AI ${idx + 1}${rec.aiBullNo ? ` · Bull ${rec.aiBullNo}` : ""}${rec.aiSetNo ? ` · Set ${rec.aiSetNo}` : ""}${rec.result ? ` · ${rec.result}` : ""}`, "cyan");
        if (rec.result === "Negative") add(rec.aiDate, "PD", `Negative pregnancy diagnosis / result in repro parity ${parity.parityNo}`, "rose");
      });

      if (conceptionDate) {
        const calvedAlready = hasCalvingAfterDate(animal, conceptionDate);
        add(conceptionDate, "Conception", `Conceived in repro parity ${parity.parityNo}${conceivedRecord?.aiBullNo ? ` · Bull ${conceivedRecord.aiBullNo}` : ""}${!calvedAlready ? ` · ECD ${parity.expectedCalvingDate || expectedCalving(conceptionDate)}` : ""}`, "green");
      }
    });

    (animal.femaleDetails?.calvingParities || []).forEach((row) => {
      getCalvingOutcomeEvents(row).forEach((event, outcomeIdx) => {
        const outcome = event.calvingOutcome || "";
        const title = outcome === "Abortion" ? "Abortion" : outcome === "Stillbirth" ? "Stillbirth" : outcome === "Premature calving" ? "Premature calving" : "Calving";
        const productive = isProductiveCalvingOutcome(event);
        const afterAbortion = outcomeIdx > 0 && getCalvingOutcomeEvents(row).slice(0, outcomeIdx).some((prev) => prev.calvingOutcome === "Abortion");
        const detailPrefix = productive ? `Calving parity ${row.parityNo} · Conception ${outcomeIdx + 1}` : `${outcome} recorded · Conception ${outcomeIdx + 1}`;
        const showCalfIdentity = isCalfIdentityOutcome(event);
        add(event.calvingDate, title, `${detailPrefix} · ${outcome}${afterAbortion ? " · after previous abortion" : ""}${showCalfIdentity && event.calfSex ? ` · ${event.calfSex}` : ""}${showCalfIdentity && event.calfTag ? ` · Calf ${event.calfTag}` : ""}`, productive ? "amber" : "rose");
      });
    });

    (animal.femaleDetails?.productionLactations || []).forEach((lactation) => {
      const records = lactation.fridayRecords || [];
      if (records.length) {
        const latest = [...records].filter((r) => r.date).pop();
        add(latest?.date, "Milk Record", `Lactation parity ${lactation.parityNo} · Latest test-day ${latest?.totalDailyYield || ""} kg`, "blue");
      }
      add(lactation.dryDate, "Dry Date", `Lactation parity ${lactation.parityNo} dried`, "slate");
    });

    const health = animal.femaleDetails?.health || {};
    (health.treatmentRecords || []).forEach((r) => add(r.treatmentDate, "Treatment", `${r.diagnosis || "Diagnosis not entered"}${r.treatmentGiven ? ` · ${r.treatmentGiven}` : ""}`, "violet"));
  } else {
    const health = animal.maleDetails?.health || {};
    (animal.maleDetails?.diseaseTests || []).forEach((r) => add(r.testDate, "Disease Test", `${r.testName || "Test"}${r.result ? ` · ${r.result}` : ""}`, "cyan"));
    (health.treatmentRecords || []).forEach((r) => add(r.treatmentDate, "Treatment", `${r.diagnosis || "Diagnosis not entered"}${r.treatmentGiven ? ` · ${r.treatmentGiven}` : ""}`, "violet"));
  }

  add(animal.exitDate, animal.status || "Archive", animal.exitReason || "Archived / exited", "rose");

  return events
    .filter((item) => parseDisplayDate(item.date))
    .sort((a, b) => parseDisplayDate(a.date).getTime() - parseDisplayDate(b.date).getTime());
}

function AnimalSummaryPanel({ animal }) {
  const summary = getAnimalSummarySnapshot(animal);
  if (!summary) return null;
  const rows = [
    ["Tag no.", summary.tagNo],
    ["Age", summary.age],
    ["Current parity", summary.currentParity],
    ["Last calving", summary.lastCalving],
    ["Last AI", summary.lastAI],
    ["Pregnancy status", summary.pregnancyStatus],
    ["Expected calving", summary.expectedCalving],
    ["Milk status", summary.milkStatus],
  ];
  return (
    <Section title="Animal Summary">
      <div className="space-y-3">
        {rows.map(([label, value]) => (
          <div key={label} className="flex items-center justify-between gap-4 rounded-2xl border border-slate-100 bg-slate-50/70 px-4 py-3 text-sm">
            <span className="font-medium text-slate-500">{label}</span>
            <span className="text-right font-semibold text-slate-950">{value || "—"}</span>
          </div>
        ))}
        <div className="rounded-2xl border border-emerald-100 bg-emerald-50/60 p-4">
          <div className="text-xs font-semibold uppercase tracking-[0.18em] text-emerald-700">Alerts</div>
          <div className="mt-3 flex flex-wrap gap-2">
            {summary.alerts.length ? summary.alerts.map((alert) => (
              <span key={alert} className="rounded-full bg-white px-3 py-1 text-xs font-semibold text-rose-700 shadow-sm">{alert}</span>
            )) : <span className="text-sm font-medium text-emerald-700">No active alert</span>}
          </div>
        </div>
      </div>
    </Section>
  );
}

function AnimalTimelinePanel({ animal }) {
  const timeline = buildAnimalTimeline(animal);
  return (
    <div className="space-y-5">
      {timeline.length ? timeline.map((item, idx) => (
        <div key={`${item.title}-${item.date}-${idx}`} className="relative ml-4 rounded-[22px] border border-slate-200/70 bg-white/95 p-4 shadow-[0_10px_24px_rgba(15,23,42,0.045)]">
          <div className="absolute -left-5 top-5 h-4 w-4 rounded-full border-4 border-white bg-emerald-500 shadow" />
          {idx < timeline.length - 1 && <div className="absolute -left-[13px] top-9 h-[calc(100%+1rem)] w-0.5 bg-slate-200" />}
          <div className="flex flex-col gap-1 sm:flex-row sm:items-center sm:justify-between">
            <div className="text-lg font-bold text-slate-950">{item.title}</div>
            <div className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-600">{item.date}</div>
          </div>
          <div className="mt-2 text-sm leading-6 text-slate-600">{item.detail}</div>
        </div>
      )) : (
        <EmptyState title="No timeline events yet" description="DOB, AI, conception, PD, calving, milk record, treatment and dry date events will appear here once entered." />
      )}
    </div>
  );
}

function computeProductionMetrics(lactation) {
  if (!lactation) return { lactationLength: 0, totalLactationMilk: 0, standardLactationMilk: 0, peakYield: 0 };
  const calvingDate = lactation.calvingDate || "";
  const dryDate = lactation.dryDate || "";
  const lactationLength = calvingDate && dryDate ? daysBetween(calvingDate, dryDate) + 1 : 0;

  if (lactation.entryMode === "Manual") {
    return {
      lactationLength,
      totalLactationMilk: Number(lactation.manualSummary.totalLactationMilk || 0),
      standardLactationMilk: Number(lactation.manualSummary.standardLactationMilk || 0),
      peakYield: Number(lactation.manualSummary.peakYield || 0),
    };
  }

  const records = [...(lactation.fridayRecords || [])].filter((r) => r.date).sort((a, b) => {
    const ad = parseDisplayDate(a.date);
    const bd = parseDisplayDate(b.date);
    if (!ad || !bd) return 0;
    return ad.getTime() - bd.getTime();
  });

  const firstFriday = firstRecordableFriday(calvingDate);
  let total = 0;
  let standard = 0;
  let peak = 0;
  let usedDays = 0;
  let standardUsed = 0;
  const hardLength = lactationLength > 0 ? lactationLength : Infinity;

  records.forEach((r, index) => {
    const milk = Number(r.totalDailyYield || 0);
    peak = Math.max(peak, milk);
    const fullBlock = index === 0 && firstFriday && r.date === firstFriday ? daysBetween(calvingDate, firstFriday) + 7 : 7;
    const applied = Math.max(0, Math.min(fullBlock, hardLength - usedDays));
    usedDays += applied;
    total += milk * applied;
    const standardApplied = Math.max(0, Math.min(applied, 305 - standardUsed));
    standardUsed += standardApplied;
    standard += milk * standardApplied;
  });

  return { lactationLength, totalLactationMilk: total, standardLactationMilk: standard, peakYield: peak };
}

function getMonthlyMilkSplit(lactation) {
  if (!lactation || lactation.entryMode !== "Friday Records") return [];
  const calvingDate = lactation.calvingDate || "";
  const dryDate = lactation.dryDate || "";
  const calving = parseDisplayDate(calvingDate);
  if (!calving) return [];

  const records = [...(lactation.fridayRecords || [])].filter((r) => r.date).sort((a, b) => {
    const ad = parseDisplayDate(a.date);
    const bd = parseDisplayDate(b.date);
    if (!ad || !bd) return 0;
    return ad.getTime() - bd.getTime();
  });
  if (!records.length) return [];

  const firstFriday = firstRecordableFriday(calvingDate);
  const lactationLength = calvingDate && dryDate ? daysBetween(calvingDate, dryDate) + 1 : Infinity;
  const monthMap = new Map();
  let usedDays = 0;

  records.forEach((r, index) => {
    const milk = Number(r.totalDailyYield || 0);
    const fullBlock = index === 0 && firstFriday && r.date === firstFriday ? daysBetween(calvingDate, firstFriday) + 7 : 7;
    const applied = Math.max(0, Math.min(fullBlock, lactationLength - usedDays));
    if (applied <= 0) return;
    const blockStart = parseDisplayDate(index === 0 && firstFriday && r.date === firstFriday ? calvingDate : r.date);
    if (!blockStart) return;

    for (let dayOffset = 0; dayOffset < applied; dayOffset += 1) {
      const d = new Date(blockStart);
      d.setDate(d.getDate() + dayOffset);
      const key = `${d.toLocaleString("en-US", { month: "long" })} ${d.getFullYear()}`;
      const existing = monthMap.get(key) || { month: key, milkYield: 0 };
      existing.milkYield += milk;
      monthMap.set(key, existing);
    }

    usedDays += applied;
  });

  return Array.from(monthMap.values()).map((row) => ({ ...row, milkYield: Number(row.milkYield.toFixed(1)) }));
}

function computeHistoryRows(animal) {
  const rows = [];

  for (let p = 1; p <= 10; p += 1) {
    const repro = getReproParityByNo(animal, p - 1);
    const calving = animal?.femaleDetails?.calvingParities?.find((c) => Number(c.parityNo) === p) || null;
    const events = getCalvingOutcomeEvents(calving).filter((event) => event.calvingDate || event.calvingOutcome || event.remarks || event.calfTag);
    const lactation = animal?.femaleDetails?.productionLactations?.find((l) => Number(l.parityNo) === p) || null;
    const prod = lactation ? computeProductionMetrics(lactation) : null;
    const aiRecords = sortAIRecords(repro?.aiRecords || []);

    if (!events.length && !repro && !lactation) continue;

    const rowsForParity = events.length ? events : [makeCalvingOutcomeEvent(1)];

    rowsForParity.forEach((event, outcomeIndex) => {
      const conceptionRecord = getConceptionRecordForOutcome(animal, p, outcomeIndex);
      const conceptionDate = conceptionRecord?.aiDate || "";
      const aiCycle = getAIServicesForConceptionCycle(animal, p, outcomeIndex);
      const firstAI = getFirstAIForConceptionCycle(animal, p, outcomeIndex);
      const bullNo = formatBullSet(conceptionRecord || aiCycle[aiCycle.length - 1] || aiRecords[aiRecords.length - 1]);
      const conceptionNo = conceptionDate ? String(outcomeIndex + 1) : "";
      const calfTag = isCalfIdentityOutcome(event) ? (event?.calfTag || "") : "";
      const metrics = computeCalvingOutcomeMetrics(animal, p, event, outcomeIndex);

      const dryDate = isCompleteProductiveCalvingEvent(event) ? (lactation?.dryDate || "") : "";
      // Dry period is shown on the next productive calving row, like SP and CI.
      // Therefore parity 1 remains blank; parity 2 shows dry date of parity 1 to calving date of parity 2, and so on.
      const previousLactation = animal?.femaleDetails?.productionLactations?.find((l) => Number(l.parityNo) === p - 1) || null;
      const previousDryDate = previousLactation?.dryDate || "";
      const dryPeriod = previousDryDate && event?.calvingDate && isCompleteProductiveCalvingEvent(event) ? String(daysBetween(previousDryDate, event.calvingDate)) : "";

      rows.push({
        parity: String(p),
        outcomeType: event?.calvingOutcome || "",
        dateCalved: event?.calvingDate || "",
        pregnancyOutcomeDate: event?.calvingDate ? `${event?.calvingOutcome || ""} / ${event.calvingDate}` : "",
        gp: metrics.gestationPeriod === null ? "" : String(metrics.gestationPeriod),
        sexOfCalf: isCalfIdentityOutcome(event) ? event?.calfSex || "" : "",
        calfTag,
        firstAI,
        aiTillConception: conceptionDate ? String(aiCycle.length || 1) : "",
        conceptionDate,
        bullNo,
        totalAI: conceptionNo,
        dryDate,
        dryPeriod,
        tlmy: isCompleteProductiveCalvingEvent(event) && lactation ? (lactation.entryMode === "Manual" ? lactation.manualSummary.totalLactationMilk || "" : prod ? String(prod.totalLactationMilk.toFixed(1)) : "") : "",
        slmy: isCompleteProductiveCalvingEvent(event) && lactation ? (lactation.entryMode === "Manual" ? lactation.manualSummary.standardLactationMilk || "" : prod ? String(prod.standardLactationMilk.toFixed(1)) : "") : "",
        ll: isCompleteProductiveCalvingEvent(event) && lactation && lactation.dryDate && prod ? String(prod.lactationLength) : "",
        py: isCompleteProductiveCalvingEvent(event) && lactation ? (lactation.entryMode === "Manual" ? lactation.manualSummary.peakYield || "" : prod ? String(prod.peakYield.toFixed(1)) : "") : "",
        sp: metrics.servicePeriod === null ? "" : String(metrics.servicePeriod),
        ci: metrics.calvingInterval === null ? "" : String(metrics.calvingInterval),
        fat: "",
        snf: "",
        ts: "",
      });
    });
  }

  return rows;
}

function getDamRecordRowsForMale(animal, animals) {
  if (!animal || animal.category !== "Male") return null;
  const damTag = String(animal?.maleDetails?.pedigree?.dam || "").trim();
  if (!damTag) return null;
  const dam = (animals || []).find((a) => a?.category === "Female" && [a?.tagNo, a?.id].map((v) => String(v || "").trim()).includes(damTag));
  if (!dam) return null;

  const historyRows = computeHistoryRows(dam);
  const afcValue = computeCalvingMetrics(dam, 1).afc;
  const byParity = (parityNo, key) => historyRows.find((row) => Number(row.parity) === parityNo)?.[key] || "";

  return {
    damTag: dam.tagNo || damTag,
    tlmy: Array.from({ length: 10 }, (_, idx) => byParity(idx + 1, "tlmy")),
    slmy: Array.from({ length: 10 }, (_, idx) => byParity(idx + 1, "slmy")),
    py: Array.from({ length: 10 }, (_, idx) => byParity(idx + 1, "py")),
    sp: Array.from({ length: 10 }, (_, idx) => byParity(idx + 1, "sp")),
    ci: Array.from({ length: 10 }, (_, idx) => byParity(idx + 1, "ci")),
    afc: afcValue == null ? "" : String(afcValue),
  };
}

function exportHistoryPdf(animal) {
  if (!animal) return;
  const full = withDefaults(animal);
  const rows = computeHistoryRows(full);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFontSize(15);
  doc.text("BUFFALO HISTORY SHEET", pageWidth / 2, 11, { align: "center" });
  doc.setFontSize(10.5);
  doc.text("ICAR-CENTRAL INSTITUTE FOR RESEARCH ON BUFFALOES", pageWidth / 2, 18, { align: "center" });
  doc.text("SUB-CAMPUS, NABHA PUNJAB 147201", pageWidth / 2, 24, { align: "center" });
  doc.setFontSize(9.2);
  doc.text(`Animal No.: ${full.tagNo || ""}`, 6, 34);
  doc.text(`Date of Birth: ${full.dob || ""}`, 50, 34);
  doc.text(`AFC (days): ${computeCalvingMetrics(full, 1).afc ?? ""}`, 98, 34);
  doc.text(`Reason for culling: ${full.femaleDetails.historyMeta.reasonForCulling || ""}`, 140, 34);
  doc.text(`Book Value: ${full.femaleDetails.historyMeta.bookValue || ""}`, 230, 34);

  autoTable(doc, {
    startY: 39,
    styles: { fontSize: 7.2, cellPadding: 0.75, overflow: "linebreak", lineWidth: 0.10, minCellHeight: 5.2, textColor: [0, 0, 0], fontStyle: "bold" },
    headStyles: { fontSize: 7.0, fillColor: [220, 245, 232], textColor: [0, 0, 0], fontStyle: "bold", cellPadding: 0.75 },
    bodyStyles: { textColor: [0, 0, 0], fontStyle: "bold" },
    theme: "grid",
    head: [["Parity", "Date of 1st A.I", "No. of AI till Conception", "Date of Conception", "Bull No./Set No. leading to Conception", "Conception No.", "Outcome Type", "Outcome Date", "GP", "Sex of Calf", "Tag No. of Calf", "Dry Date", "Dry Period", "TLMY", "SLMY", "LL", "PY", "SP", "CI", "Fat %", "SNF %", "TS %"]],
    body: rows.map((r) => [r.parity, r.firstAI, r.aiTillConception, r.conceptionDate, r.bullNo, r.totalAI, r.outcomeType, r.dateCalved, r.gp, r.sexOfCalf, r.calfTag, r.dryDate, r.dryPeriod, r.tlmy, r.slmy, r.ll, r.py, r.sp, r.ci, r.fat, r.snf, r.ts]),
    margin: { left: 4, right: 4 },
    tableWidth: "auto",
    columnStyles: {
      0: { cellWidth: 8 },
      1: { cellWidth: 15 },
      2: { cellWidth: 18 },
      3: { cellWidth: 15 },
      4: { cellWidth: 24 },
      5: { cellWidth: 13 },
      6: { cellWidth: 15 },
      7: { cellWidth: 15 },
      8: { cellWidth: 8 },
      9: { cellWidth: 12 },
      10: { cellWidth: 16 },
      11: { cellWidth: 14 },
      12: { cellWidth: 13 },
      13: { cellWidth: 11 },
      14: { cellWidth: 11 },
      15: { cellWidth: 8 },
      16: { cellWidth: 8 },
      17: { cellWidth: 8 },
      18: { cellWidth: 8 },
      19: { cellWidth: 8 },
      20: { cellWidth: 8 },
      21: { cellWidth: 8 },
    },
  });

  doc.save(`buffalo-history-sheet-${full.tagNo || "animal"}.pdf`);
}

function exportBreedingBullHistoryPdf(animal, options = {}) {
  if (!animal) return;
  const full = withDefaults(animal);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 10;
  const lineGap = 5;

  const diseaseRows = (full.maleDetails?.diseaseTests || []).map((row) => [
    row?.testDate || "",
    row?.testName || "",
    row?.result || "",
    row?.remarks || "",
  ]);

  const femaleProgenies = (options.femaleProgenies || []).map((d) => [
    d?.tagNo || "",
    d?.dob || "",
  ]);

  const daughterProductionRows = (options.femaleProgenies || []).map((d) => {
    const lactations = d?.femaleDetails?.productionLactations || [];
    const latest = lactations.length ? lactations[lactations.length - 1] : null;
    const metrics = latest ? computeProductionMetrics(latest) : { totalLactationMilk: 0, standardLactationMilk: 0, peakYield: 0 };
    return [
      d?.tagNo || "",
      getFemaleLifecycle(d) || "",
      String((d?.femaleDetails?.calvingParities || []).filter((p) => p?.calvingDate).length || ""),
      latest ? metrics.totalLactationMilk.toFixed(1) : "",
      latest ? metrics.standardLactationMilk.toFixed(1) : "",
      latest ? metrics.peakYield.toFixed(1) : "",
    ];
  });

  const daughterReproductionRows = (options.femaleProgenies || []).map((d) => {
    const calvingParities = (d?.femaleDetails?.calvingParities || []).filter((p) => p?.calvingDate).sort((a, b) => Number(a?.parityNo) - Number(b?.parityNo));
    const afc = computeCalvingMetrics(d, 1).afc ?? "";
    const servicePeriods = calvingParities.map((cp) => computeCalvingMetrics(d, cp.parityNo).servicePeriod).filter((v) => v !== null && v !== undefined && v !== "").join(", ");
    const calvingIntervals = calvingParities.map((cp) => computeCalvingMetrics(d, cp.parityNo).calvingInterval).filter((v) => v !== null && v !== undefined && v !== "").join(", ");
    return [d?.tagNo || "", String(afc), servicePeriods, calvingIntervals];
  });

  const bullAIStats = options.selectedBullAIStats || {};
  const bullAIEvents = (options.selectedBullAIEvents || []).map((row) => [
    row?.femaleTag || "",
    row?.parityNo || "",
    row?.bullAiNo || "",
    row?.aiDate || "",
    row?.setNo || "—",
    row?.result || "",
    row?.conceptionDate || "—",
  ]);

  const addTitleBlock = () => {
    doc.setFontSize(14);
    doc.text("BREEDING BULL HISTORY SHEET", pageWidth / 2, 12, { align: "center" });
    doc.setFontSize(10);
    doc.text("ICAR-CENTRAL INSTITUTE FOR RESEARCH ON BUFFALOES", pageWidth / 2, 18, { align: "center" });
    doc.text("SUB-CAMPUS, NABHA PUNJAB 147201", pageWidth / 2, 23, { align: "center" });
    doc.setFontSize(9);
    doc.text(`Bull No.: ${full.tagNo || ""}`, marginX, 31);
    doc.text(`Date of Birth: ${full.dob || ""}`, 70, 31);
    doc.text(`Breeding Set: ${full.breedingSet || ""}`, 126, 31);
    doc.text(`Sire: ${full.maleDetails?.pedigree?.sire || ""}`, 180, 31);
    doc.text(`Dam: ${full.maleDetails?.pedigree?.dam || ""}`, 235, 31);

    doc.text(`Female progenies: ${options.femaleProgenies?.length || 0}`, marginX, 36);
    doc.text(`Male progenies: ${options.maleProgenies?.length || 0}`, 70, 36);
    doc.text(`Book Value: ${full.maleDetails?.historyMeta?.bookValue || ""}`, 126, 36);
    doc.text(`Remarks: ${full.maleDetails?.historyMeta?.remarks || ""}`, 180, 36, { maxWidth: 100 });
  };

  const ensureSpace = (needed = 18) => {
    const currentY = (doc.lastAutoTable?.finalY || 40) + 6;
    if (currentY + needed > pageHeight - 10) {
      doc.addPage();
      return 14;
    }
    return currentY;
  };

  const addSection = (title, head, body, options = {}) => {
    const safeBody = body && body.length ? body : [["—"]];
    const safeHead = head && head.length ? head : [""];
    const startY = ensureSpace(options.estimatedHeight || 24);
    doc.setFontSize(11);
    doc.text(title, marginX, startY);
    autoTable(doc, {
      startY: startY + 3,
      head: [safeHead],
      body: safeBody,
      theme: "grid",
      styles: { fontSize: options.fontSize || 8, cellPadding: 1.6, overflow: "linebreak" },
      headStyles: { fillColor: [220, 245, 232], textColor: 20, fontStyle: "bold" },
      margin: { left: marginX, right: marginX },
      ...options.tableOptions,
    });
  };

  addTitleBlock();

  addSection(
    "Disease Testing",
    ["Disease test date", "Disease/test", "Result", "Remarks"],
    diseaseRows.length ? diseaseRows : [["", "", "", ""]],
    { estimatedHeight: Math.max(24, diseaseRows.length * lineGap + 16) }
  );

  addSection(
    "AI Summary",
    ["No. of AI", "No. conceived", "Negative", "Pending", "CR %", "Conceived in first attempt", "FSCR %"],
    [[
      bullAIStats.totalAI ?? 0,
      bullAIStats.conceived ?? 0,
      bullAIStats.negative ?? 0,
      bullAIStats.pending ?? 0,
      bullAIStats.crPct ?? "0.00",
      bullAIStats.firstAttemptConceived ?? 0,
      bullAIStats.firstServiceConceptionRate ?? "0.00",
    ]],
    { estimatedHeight: 24 }
  );

  addSection(
    "AI Records",
    ["Female Tag", "Conception Date"],
    bullAIEvents.length ? bullAIEvents.map((row) => [row[0], row[6]]) : [["", ""]],
    { estimatedHeight: Math.max(28, bullAIEvents.length * 5 + 16), fontSize: 8 }
  );

  addSection(
    "Female Progenies",
    ["Tag No.", "DOB"],
    femaleProgenies.length ? femaleProgenies : [["", ""]],
    { estimatedHeight: Math.max(24, femaleProgenies.length * lineGap + 16) }
  );

  addSection(
    "Performance of Daughters · Production",
    ["Daughter", "Current category", "Parity count", "Latest TLMY", "Latest SLMY", "Peak yield"],
    daughterProductionRows.length ? daughterProductionRows : [["", "", "", "", "", ""]],
    { estimatedHeight: Math.max(24, daughterProductionRows.length * lineGap + 16) }
  );

  addSection(
    "Performance of Daughters · Reproduction",
    ["Daughter", "AFC (days)", "Service periods (days)", "Calving Intervals (days)"],
    daughterReproductionRows.length ? daughterReproductionRows : [["", "", "", ""]],
    { estimatedHeight: Math.max(24, daughterReproductionRows.length * lineGap + 16) }
  );

  doc.save(`breeding-bull-history-sheet-${full.tagNo || "animal"}.pdf`);
}

function AnimalPhotographsPanel({ photos = [], onUpload, onRemove }) {
  const remaining = Math.max(0, 10 - photos.length);
  return (
    <div className="space-y-5">
      <div className="rounded-[28px] border-2 border-dashed border-emerald-400 bg-white/75 p-6 text-center">
        <div className="mx-auto mb-3 flex h-14 w-14 items-center justify-center rounded-2xl bg-emerald-50 text-3xl text-emerald-700">↑</div>
        <div className="text-lg font-semibold text-slate-800">Upload animal photographs</div>
        <div className="mt-1 text-sm text-slate-600">JPEG, PNG, HEIC only · maximum 10 pictures per animal</div>
        <label className={`mx-auto mt-4 inline-flex cursor-pointer items-center justify-center rounded-2xl px-5 py-3 font-semibold shadow-sm ${remaining ? "bg-emerald-600 text-white hover:bg-emerald-700" : "bg-slate-200 text-slate-500"}`}>
          {remaining ? `Select photographs (${remaining} slot${remaining === 1 ? "" : "s"} left)` : "Maximum 10 saved"}
          <input type="file" className="hidden" multiple accept="image/jpeg,image/png,image/heic,image/heif,.jpg,.jpeg,.png,.heic,.heif" onChange={onUpload} disabled={!remaining} />
        </label>
      </div>

      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-5">
        {photos.map((photo, idx) => {
          const isHeic = String(photo.type || photo.name || "").toLowerCase().includes("heic") || String(photo.name || "").toLowerCase().endsWith(".heic");
          return (
            <div key={photo.id || `${photo.name}-${idx}`} className="overflow-hidden rounded-3xl border border-emerald-100 bg-white shadow-sm">
              <div className="flex aspect-square items-center justify-center bg-slate-50">
                {isHeic ? (
                  <div className="px-3 text-center text-sm font-semibold text-slate-500">HEIC photograph saved</div>
                ) : (
                  <img src={photo.dataUrl} alt={photo.name || `Animal photograph ${idx + 1}`} className="h-full w-full object-cover" />
                )}
              </div>
              <div className="space-y-2 p-3">
                <div className="truncate text-xs font-medium text-slate-700" title={photo.name}>{photo.name || `Photograph ${idx + 1}`}</div>
                <button className="w-full rounded-xl border border-rose-200 bg-white px-3 py-2 text-sm font-medium text-rose-700 hover:bg-rose-50" onClick={() => onRemove(photo.id)}>Remove</button>
              </div>
            </div>
          );
        })}
        {Array.from({ length: remaining }).map((_, idx) => (
          <div key={`empty-photo-${idx}`} className="flex aspect-square items-center justify-center rounded-3xl border border-dashed border-slate-200 bg-white/55 text-sm text-slate-400">Empty photo slot</div>
        ))}
      </div>
    </div>
  );
}

function Section({ title, children }) {
  return (
    <section className="ds-card group relative overflow-hidden rounded-[28px] border border-slate-200/70 bg-white/94 p-6 shadow-[0_14px_34px_rgba(15,23,42,0.065)] backdrop-blur-xl transition duration-300 hover:shadow-[0_18px_44px_rgba(15,23,42,0.09)] sm:p-7">
      <div className="pointer-events-none absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-emerald-500 via-cyan-400 to-transparent opacity-80" />
      <div className="relative mb-5 border-b border-slate-100 pb-4">
        <h2 className="text-2xl font-bold tracking-tight text-slate-950">{title}</h2>
      </div>
      <div className="relative">{children}</div>
    </section>
  );
}

function Grid({ children }) {
  return <div className="grid grid-cols-1 gap-4 md:grid-cols-3">{children}</div>;
}

function TextField({ label, value, onChange, type = "text", readOnly = false, placeholder = "" }) {
  return (
    <label className="flex flex-col gap-1.5 text-sm">
      <span className="font-medium text-slate-600">{label}</span>
      <input
        className="rounded-2xl border border-slate-200 bg-white/90 px-4 py-3 text-slate-800 shadow-sm transition duration-200 focus:border-emerald-400 focus:outline-none focus:ring-4 focus:ring-emerald-100"
        type={type}
        value={value}
        readOnly={readOnly}
        placeholder={placeholder}
        onChange={readOnly ? undefined : (e) => onChange(e.target.value)}
      />
    </label>
  );
}

function DateField({ label, value, onChange, readOnly = false }) {
  return <TextField label={label} value={value} onChange={(v) => onChange(normalizeDisplayDate(v))} readOnly={readOnly} placeholder="dd/mm/yyyy" />;
}

function FridayDatePicker({ label, value, onChange }) {
  const selected = parseDisplayDate(value) || new Date();
  const [open, setOpen] = useState(false);
  const [viewDate, setViewDate] = useState(() => new Date(selected.getFullYear(), selected.getMonth(), 1));

  const monthName = viewDate.toLocaleString("en-US", { month: "long", year: "numeric" });
  const first = new Date(viewDate.getFullYear(), viewDate.getMonth(), 1);
  const gridStart = new Date(first);
  gridStart.setDate(first.getDate() - first.getDay());
  const days = Array.from({ length: 42 }, (_, idx) => {
    const d = new Date(gridStart);
    d.setDate(gridStart.getDate() + idx);
    return d;
  });

  function shiftMonth(delta) {
    setViewDate((prev) => new Date(prev.getFullYear(), prev.getMonth() + delta, 1));
  }

  function pickFriday(day) {
    if (day.getDay() !== 5) return;
    onChange(formatDateDisplay(day));
    setOpen(false);
  }

  return (
    <div className="relative flex flex-col gap-1.5 text-sm">
      <span className="font-medium text-slate-600">{label}</span>
      <button
        type="button"
        className="flex items-center justify-between gap-2 rounded-2xl border border-slate-200 bg-white/90 px-4 py-3 text-left text-slate-800 shadow-sm transition duration-200 focus:border-emerald-400 focus:outline-none focus:ring-4 focus:ring-emerald-100"
        onClick={() => setOpen((prev) => !prev)}
      >
        <span>{value || "Select Friday"}</span>
        <span className="text-lg" aria-hidden="true">📅</span>
      </button>
      {open && (
        <div className="absolute right-0 top-full z-50 mt-2 w-[275px] rounded-[22px] border border-emerald-100 bg-white p-3 shadow-[0_20px_50px_rgba(15,23,42,0.18)]">
          <div className="mb-2 flex items-center justify-between">
            <button type="button" className="rounded-lg border border-emerald-200 px-2.5 py-1 text-sm font-bold text-emerald-800 hover:bg-emerald-50" onClick={() => shiftMonth(-1)}>‹</button>
            <div className="text-sm font-bold text-emerald-950">{monthName}</div>
            <button type="button" className="rounded-lg border border-emerald-200 px-2.5 py-1 text-sm font-bold text-emerald-800 hover:bg-emerald-50" onClick={() => shiftMonth(1)}>›</button>
          </div>
          <div className="grid grid-cols-7 gap-1 text-center text-[11px] font-bold uppercase tracking-wide text-slate-500">
            {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d) => <div key={d}>{d}</div>)}
          </div>
          <div className="mt-1.5 grid grid-cols-7 gap-1">
            {days.map((day) => {
              const isCurrentMonth = day.getMonth() === viewDate.getMonth();
              const isFriday = day.getDay() === 5;
              const isSelected = formatDateDisplay(day) === value;
              return (
                <button
                  key={day.toISOString()}
                  type="button"
                  disabled={!isFriday}
                  onClick={() => pickFriday(day)}
                  className={`h-9 rounded-xl text-sm font-semibold transition ${
                    isFriday
                      ? isSelected
                        ? "bg-emerald-600 text-white shadow-sm"
                        : "border border-emerald-200 bg-emerald-50 text-emerald-900 hover:bg-emerald-100"
                      : "cursor-not-allowed bg-slate-50 text-slate-300"
                  } ${!isCurrentMonth ? "opacity-40" : ""}`}
                  title={isFriday ? "Select Friday" : "Only Fridays are selectable"}
                >
                  {day.getDate()}
                </button>
              );
            })}
          </div>
          <div className="mt-2 rounded-xl bg-emerald-50 px-2.5 py-1.5 text-[11px] text-emerald-800">Only Friday dates are active for test-day entry.</div>
        </div>
      )}
    </div>
  );
}

function SelectField({ label, value, onChange, options, disabled = false }) {
  return (
    <label className="flex flex-col gap-1.5 text-sm">
      <span className="font-medium text-slate-600">{label}</span>
      <select
        className="rounded-2xl border border-slate-200 bg-white/90 px-4 py-3 text-slate-800 shadow-sm transition duration-200 focus:border-emerald-400 focus:outline-none focus:ring-4 focus:ring-emerald-100"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
      >
        {options.map((o) => {
          const optValue = typeof o === "object" && o !== null ? o.value : o;
          const optLabel = typeof o === "object" && o !== null ? o.label : o;
          return (
            <option key={String(optValue)} value={optValue}>
              {optLabel || "—"}
            </option>
          );
        })}
      </select>
    </label>
  );
}

function AreaField({ label, value, onChange, rows = 4, readOnly = false }) {
  return (
    <label className="flex flex-col gap-1 text-sm md:col-span-3">
      <span>{label}</span>
      <textarea className="rounded-xl border border-emerald-200 px-3 py-2 focus:border-emerald-500 focus:outline-none" rows={rows} value={value} readOnly={readOnly} onChange={readOnly ? undefined : (e) => onChange(e.target.value)} />
    </label>
  );
}

function StatCard({ title, value, onClick, hint }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`group relative flex min-h-[145px] w-full flex-col justify-between overflow-hidden rounded-[24px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] backdrop-blur-xl transition duration-300 hover:-translate-y-0.5 hover:shadow-[0_18px_38px_rgba(15,23,42,0.10)] ${onClick ? "cursor-pointer" : "cursor-default"}`}
    >
      <div className="pointer-events-none absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-emerald-500 via-cyan-400 to-transparent opacity-80" />
      <div className="relative min-h-[2.75rem] text-[11px] font-semibold uppercase leading-5 tracking-[0.16em] text-slate-500">{title}</div>
      <div className="relative mt-4 text-3xl font-bold tracking-tight text-slate-950 sm:text-[2rem]">{value}</div>
      {hint ? <div className="relative mt-3 text-xs font-semibold text-emerald-700">{hint}</div> : null}
    </button>
  );
}

function EmptyState({ title, description, actionLabel, onAction }) {
  return (
    <div className="flex min-h-[190px] flex-col items-center justify-center rounded-[24px] border border-dashed border-slate-200/80 bg-white/80 p-7 text-center shadow-[0_10px_24px_rgba(15,23,42,0.045)]">
      <div className="grid h-14 w-14 place-items-center rounded-2xl bg-slate-950 text-xl text-white shadow-lg ring-8 ring-slate-100">◎</div>
      <div className="mt-5 text-lg font-bold tracking-tight text-slate-950">{title}</div>
      <div className="mt-2 max-w-sm text-sm leading-6 text-slate-500">{description}</div>
      {actionLabel && onAction ? (
        <button
          className="mt-5 rounded-2xl bg-emerald-600 px-4 py-3 text-sm font-semibold text-white shadow-lg transition duration-200 hover:-translate-y-0.5 hover:bg-emerald-700"
          onClick={onAction}
        >
          {actionLabel}
        </button>
      ) : null}
    </div>
  );
}

function ProgressMetric({ label, value, tone = "emerald" }) {
  const safe = Math.max(0, Math.min(100, Number(value) || 0));
  const tones = {
    emerald: "from-emerald-400 to-emerald-500",
    cyan: "from-cyan-400 to-cyan-500",
    amber: "from-amber-400 to-orange-400",
    violet: "from-violet-400 to-indigo-400",
  };
  return (
    <div className="rounded-[24px] border border-white/70 bg-white/88 p-5 shadow-[0_14px_30px_rgba(15,23,42,0.07)]">
      <div className="flex items-center justify-between">
        <div className="text-sm font-medium text-slate-600">{label}</div>
        <div className="text-sm font-semibold text-slate-900">{safe.toFixed(0)}%</div>
      </div>
      <div className="mt-4 h-2 overflow-hidden rounded-full bg-slate-100">
        <div className={`h-full rounded-full bg-gradient-to-r ${tones[tone] || tones.emerald}`} style={{ width: `${safe}%` }} />
      </div>
    </div>
  );
}

function SparkBars({ title, data, labels = [] }) {
  const max = Math.max(1, ...(data || [0]));
  return (
    <div className="rounded-[28px] border border-white/70 bg-white/90 p-5 shadow-[0_18px_45px_rgba(15,23,42,0.08)] backdrop-blur-xl">
      <div className="mb-4 flex items-center justify-between">
        <div className="text-lg font-semibold tracking-tight text-slate-900">{title}</div>
        <div className="text-xs uppercase tracking-[0.18em] text-slate-400">Live</div>
      </div>
      <div className="flex h-48 items-end gap-3">
        {(data || []).map((value, idx) => (
          <div key={`${title}-${idx}`} className="flex flex-1 flex-col items-center gap-2">
            <div className="flex h-40 w-full items-end">
              <div
                className="w-full rounded-t-2xl bg-gradient-to-t from-emerald-500 to-cyan-400 shadow-md"
                style={{ height: `${Math.max(8, (Number(value || 0) / max) * 100)}%` }}
              />
            </div>
            <div className="text-xs font-medium text-slate-500">{labels[idx] || idx + 1}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function MiniInsight({ title, subtitle, value, tone = "emerald" }) {
  const accents = {
    emerald: "bg-emerald-50 text-emerald-700",
    cyan: "bg-cyan-50 text-cyan-700",
    amber: "bg-amber-50 text-amber-700",
    violet: "bg-violet-50 text-violet-700",
  };
  return (
    <div className="rounded-[24px] border border-white/70 bg-white/92 p-5 shadow-[0_14px_30px_rgba(15,23,42,0.07)] transition duration-300 hover:-translate-y-0.5 hover:shadow-[0_18px_40px_rgba(15,23,42,0.09)]">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="text-sm font-medium text-slate-500">{title}</div>
          <div className="mt-1 text-2xl font-semibold tracking-tight text-slate-900">{value}</div>
        </div>
        <div className={`rounded-2xl px-3 py-2 text-xs font-semibold ${accents[tone] || accents.emerald}`}>{subtitle}</div>
      </div>
    </div>
  );
}



function getCurrentFemaleReproParity(animal) {
  if (!animal || animal.category !== "Female") return null;
  const normalCalvings = (animal?.femaleDetails?.calvingParities || [])
    .filter((c) => c.calvingDate && (c.calvingOutcome || "Normal calving") === "Normal calving")
    .map((c) => Number(c.parityNo || 0));
  const currentParityNo = normalCalvings.length ? Math.max(...normalCalvings) : 0;
  return getReproParityByNo(animal, currentParityNo) || getReproParityByNo(animal, 0) || null;
}

function getCurrentFemaleReproStatus(animal) {
  const parity = getCurrentFemaleReproParity(animal);
  if (!parity) return "Not inseminated";
  const aiRecords = sortAIRecords(parity.aiRecords || []);
  if (getConceivedAIRecord(parity)) return "Conceived";
  if (!aiRecords.length) return "Not inseminated";
  const last = aiRecords[aiRecords.length - 1];
  if ((last?.result || "") === "Pending") return "Pending";
  return "Inseminated but Not pregnant";
}

function getDaysInMilkBucket(animal) {
  const lifecycle = getFemaleLifecycle(animal);
  if (lifecycle !== "Milk") return null;
  const lastCalving = getLastNormalCalving(animal);
  if (!lastCalving?.calvingDate) return null;
  const dim = daysBetween(lastCalving.calvingDate, formatDateDisplay(new Date()));
  if (dim <= 60) return "<60 days";
  if (dim <= 90) return "61-90 days";
  if (dim <= 120) return "91-120 days";
  return ">120 days";
}

function getHeiferAgeBucket(animal) {
  const lifecycle = getFemaleLifecycle(animal);
  if (!["Heifer", "Colostrum-Heifer"].includes(lifecycle)) return null;
  const dob = parseDisplayDate(animal?.dob || "");
  if (!dob) return null;
  const now = new Date();
  const months = (now.getFullYear() - dob.getFullYear()) * 12 + (now.getMonth() - dob.getMonth()) - (now.getDate() < dob.getDate() ? 1 : 0);
  if (months < 24) return "<24 months";
  if (months < 30) return "24-30 months";
  if (months < 36) return "30-36 months";
  if (months < 42) return "36-42 months";
  if (months < 48) return "42-48 months";
  return ">48 months";
}

function makeStatusRowMap(columnKeys, includeColostrum = false) {
  const statuses = ["Conceived", "Pending", "Inseminated but Not pregnant", "Not inseminated", ...(includeColostrum ? ["Colostrum"] : [])];
  const out = {};
  statuses.forEach((s) => {
    out[s] = Object.fromEntries([...columnKeys, "Total"].map((k) => [k, 0]));
  });
  out.Total = Object.fromEntries([...columnKeys, "Total"].map((k) => [k, 0]));
  return out;
}

function buildMilkStatusSummary(animals) {
  const cols = ["<60 days", "61-90 days", "91-120 days", ">120 days"];
  const out = makeStatusRowMap(cols);
  (animals || []).filter((a) => a.category === "Female" && !isArchivedAnimal(a)).forEach((animal) => {
    const bucket = getDaysInMilkBucket(animal);
    if (!bucket) return;
    const status = getCurrentFemaleReproStatus(animal);
    out[status][bucket] += 1;
    out[status].Total += 1;
    out.Total[bucket] += 1;
    out.Total.Total += 1;
  });
  return { columns: cols, rows: out };
}

function buildDryStatusSummary(animals) {
  const statuses = ["Conceived", "Pending", "Inseminated but Not pregnant", "Not inseminated", "Colostrum"];
  const out = Object.fromEntries([...statuses, "Total"].map((s) => [s, 0]));
  (animals || []).filter((a) => a.category === "Female" && !isArchivedAnimal(a) && countsAsDryOnDate(a, formatDateDisplay(new Date()))).forEach((animal) => {
    const status = getFemaleLifecycle(animal) === "Colostrum" ? "Colostrum" : getCurrentFemaleReproStatus(animal);
    out[status] += 1;
    out.Total += 1;
  });
  return out;
}

function buildHeiferStatusSummary(animals) {
  const cols = ["<24 months", "24-30 months", "30-36 months", "36-42 months", "42-48 months", ">48 months"];
  const out = makeStatusRowMap(cols, true);
  (animals || []).filter((a) => a.category === "Female" && !isArchivedAnimal(a)).forEach((animal) => {
    const bucket = getHeiferAgeBucket(animal);
    if (!bucket) return;
    const status = getFemaleLifecycle(animal) === "Colostrum-Heifer" ? "Colostrum" : getCurrentFemaleReproStatus(animal);
    out[status][bucket] += 1;
    out[status].Total += 1;
    out.Total[bucket] += 1;
    out.Total.Total += 1;
  });
  return { columns: cols, rows: out };
}

function buildHerdSnapshotRows(animals) {
  const active = (animals || []).filter((a) => !isArchivedAnimal(a));
  const females = active.filter((a) => a.category === "Female");
  const males = active.filter((a) => a.category === "Male");
  return [
    { label: "Total active herd", value: active.length },
    { label: "Total females", value: females.length },
    { label: "Heifers", value: females.filter((a) => countsAsHeiferOnDate(a, formatDateDisplay(new Date()))).length },
    { label: "Colostrum-Heifers", value: females.filter((a) => getFemaleLifecycle(a) === "Colostrum-Heifer").length },
    { label: "Colostrum adults", value: females.filter((a) => getFemaleLifecycle(a) === "Colostrum").length },
    { label: "Buffaloes in milk", value: females.filter((a) => getFemaleLifecycle(a) === "Milk").length },
    { label: "Dry buffaloes", value: females.filter((a) => countsAsDryOnDate(a, formatDateDisplay(new Date()))).length },
    { label: "Total males", value: males.length },
    { label: "Breeding bulls", value: males.filter((a) => a.isBreedingBull === "Yes").length },
    { label: "Teasers / other males", value: males.filter((a) => a.isBreedingBull !== "Yes").length },
  ];
}

function DashboardMatrixTable({ columns, rows, includeColostrum = false }) {
  const rowOrder = ["Conceived", "Pending", "Inseminated but Not pregnant", "Not inseminated", ...(includeColostrum ? ["Colostrum"] : []), "Total"];
  return (
    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
      <table className="min-w-full text-sm">
        <thead className="bg-slate-50/90">
          <tr>
            <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Reproductive status</th>
            {columns.map((c) => <th key={c} className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">{c}</th>)}
            <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Total</th>
          </tr>
        </thead>
        <tbody>
          {rowOrder.map((label) => (
            <tr key={label}>
              <td className="whitespace-nowrap px-4 py-3 text-slate-700 font-medium">{label}</td>
              {columns.map((c) => <td key={c} className="whitespace-nowrap px-4 py-3 text-slate-700">{rows[label]?.[c] ?? 0}</td>)}
              <td className="whitespace-nowrap px-4 py-3 text-slate-700">{rows[label]?.Total ?? 0}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}


function getMonthEndDate(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function addDateOffset(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function isDateWithinMonth(dateStr, startDate, endDate) {
  const dt = parseDisplayDate(dateStr || "");
  return Boolean(dt && dt.getTime() >= startDate.getTime() && dt.getTime() <= endDate.getTime());
}

function getMovementExitType(animal) {
  const reason = String(animal?.exitReason || "").toLowerCase();
  const status = String(animal?.status || "").toLowerCase();
  const text = `${reason} ${status}`;
  if (text.includes("death") || text.includes("dead") || text.includes("mortality") || text.includes("died")) return "D";
  if (text.includes("transfer")) return "TOut";
  if (text.includes("sold") || text.includes("sale") || text.includes("cull") || text.includes("culled") || text.includes("disposal") || text.includes("disposed")) return "S";
  return isArchivedAnimal(animal) ? "S" : "";
}

function getMovementCategory(animal, date) {
  if (!animal || !date) return "";
  const dateStr = formatDateDisplay(date);
  const ageDays = ageDaysOnDate(animal, date);
  const ageMonths = ageMonthsOnDate(animal, date);

  if (animal.category === "Female") {
    if (ageDays !== null && ageDays < 90) return "F:Below 3 months";
    if (ageMonths !== null && ageMonths >= 3 && ageMonths < 12) return "F:3-12 months";
    if (ageMonths !== null && ageMonths >= 12 && ageMonths < 24) return "F:1-2 years";
    const lifecycle = getFemaleLifecycleOnDate(animal, dateStr);
    if (lifecycle === "Milk") return "F:Buffaloes in Milk";
    if (lifecycle === "Dry" || lifecycle === "Colostrum") return "F:Buffaloes Dry P/NP";
    return "F:Above 2 years";
  }

  if (animal.category === "Male") {
    if (ageDays !== null && ageDays < 90) return "M:Below 3 months";
    if (ageMonths !== null && ageMonths >= 3 && ageMonths < 12) return "M:3-12 months";
    if (ageMonths !== null && ageMonths >= 12 && ageMonths < 24) return "M:1-2 years";
    if (animal.isBreedingBull === "Yes") return "M:Breeding bulls";
    return "M:Bullocks/Teasers/others";
  }

  return "";
}

function createMovementRows(section) {
  const female = [
    "Below 3 months",
    "3-12 months",
    "1-2 years",
    "Above 2 years",
    "Buffaloes in Milk",
    "Buffaloes Dry P/NP",
  ];
  const male = [
    "Below 3 months",
    "3-12 months",
    "1-2 years",
    "Above 2 years",
    "Breeding bulls",
    "Bullocks/Teasers/others",
  ];
  const labels = section === "F" ? female : male;
  return labels.map((label) => ({
    key: `${section}:${label}`,
    label,
    ob: 0,
    b: 0,
    p: 0,
    tin: 0,
    d: 0,
    s: 0,
    tout: 0,
    cb: 0,
  }));
}

function getAnimalEffectiveEntryDate(animal) {
  if (animal?.entrySource === "Purchased") return parseDisplayDate(animal.entryDate || animal.dob || "");
  return parseDisplayDate(animal?.dob || "");
}

function isAnimalPurchasedDuringMonth(animal, startDate, endDate) {
  if (animal?.entrySource !== "Purchased") return false;
  const entry = getAnimalEffectiveEntryDate(animal);
  return Boolean(entry && entry.getTime() >= startDate.getTime() && entry.getTime() <= endDate.getTime());
}

function buildMonthlyHerdMovementReport(animals, fyStartYear, monthLabel) {
  const startDate = getFinancialMonthDate(fyStartYear, monthLabel);
  const endDate = getMonthEndDate(startDate);
  const beforeStart = addDateOffset(startDate, -1);

  const rows = [...createMovementRows("F"), ...createMovementRows("M")];
  const byKey = Object.fromEntries(rows.map((row) => [row.key, row]));

  function add(key, field, value = 1) {
    if (byKey[key]) byKey[key][field] += value;
  }

  (animals || []).forEach((animal) => {
    const dob = parseDisplayDate(animal?.dob || "");
    const entryDate = getAnimalEffectiveEntryDate(animal);
    const exitDate = parseDisplayDate(animal?.exitDate || "");
    const purchasedDuringMonth = isAnimalPurchasedDuringMonth(animal, startDate, endDate);
    const bornDuringMonth = !purchasedDuringMonth && dob && dob.getTime() >= startDate.getTime() && dob.getTime() <= endDate.getTime();
    const exitedDuringMonth = exitDate && exitDate.getTime() >= startDate.getTime() && exitDate.getTime() <= endDate.getTime();

    const presentBeforeStart = entryDate && entryDate.getTime() <= beforeStart.getTime() && (!exitDate || exitDate.getTime() > beforeStart.getTime());
    const presentEnd = entryDate && entryDate.getTime() <= endDate.getTime() && (!exitDate || exitDate.getTime() > endDate.getTime());

    const startCategory = presentBeforeStart ? getMovementCategory(animal, beforeStart) : "";
    const endCategory = presentEnd ? getMovementCategory(animal, endDate) : "";

    if (startCategory) add(startCategory, "ob");
    // CB is reconciled below as OB + additions - deductions, so it remains understandable even before the month is formally closed.

    if (bornDuringMonth) {
      const birthCategory = getMovementCategory(animal, dob);
      add(birthCategory, "b");
    }

    if (purchasedDuringMonth) {
      const purchaseCategory = getMovementCategory(animal, entryDate);
      add(purchaseCategory, "p");
    }

    if (exitedDuringMonth) {
      const exitReferenceDate = addDateOffset(exitDate, -1);
      const exitCategory = getMovementCategory(animal, exitReferenceDate) || getMovementCategory(animal, exitDate);
      const type = getMovementExitType(animal);
      if (type === "D") add(exitCategory, "d");
      else if (type === "TOut") add(exitCategory, "tout");
      else add(exitCategory, "s");
    }

    // Transfers due to age/lactation/status change.
    // Exclude births and final exits; those are already handled in B/D/S.
    if (presentBeforeStart && presentEnd && !bornDuringMonth && !exitedDuringMonth && startCategory && endCategory && startCategory !== endCategory) {
      add(startCategory, "tout");
      add(endCategory, "tin");
    }
  });

  rows.forEach((row) => {
    row.cb = Number(row.ob || 0) + Number(row.b || 0) + Number(row.p || 0) + Number(row.tin || 0) - Number(row.d || 0) - Number(row.s || 0) - Number(row.tout || 0);
  });

  const femaleRows = rows.filter((row) => row.key.startsWith("F:"));
  const maleRows = rows.filter((row) => row.key.startsWith("M:"));

  function totalRow(label, selectedRows) {
    return selectedRows.reduce((acc, row) => {
      ["ob", "b", "p", "tin", "d", "s", "tout", "cb"].forEach((field) => { acc[field] += Number(row[field] || 0); });
      return acc;
    }, { key: label, label, ob: 0, b: 0, p: 0, tin: 0, d: 0, s: 0, tout: 0, cb: 0 });
  }

  const femaleTotal = totalRow("Female Sub Total", femaleRows);
  const maleTotal = totalRow("Male Sub Total", maleRows);
  const grandTotal = totalRow("Grand Total", [femaleTotal, maleTotal]);

  return { startDate, endDate, femaleRows, maleRows, femaleTotal, maleTotal, grandTotal };
}

function HerdMovementTable({ report }) {
  const sections = [
    { title: "Female", rows: report.femaleRows, total: report.femaleTotal },
    { title: "Male", rows: report.maleRows, total: report.maleTotal },
  ];

  function renderRow(row, idx, isTotal = false) {
    return (
      <tr key={`${row.key}-${idx}`} className={isTotal ? "bg-emerald-50/60 font-semibold" : "hover:bg-emerald-50/40"}>
        <td className="whitespace-nowrap px-3 py-3 text-slate-700">{isTotal ? "" : idx + 1}</td>
        <td className="whitespace-nowrap px-3 py-3 text-slate-700 font-medium">{row.label}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.ob || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.b || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.p || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.tin || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.d || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.s || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.tout || "-"}</td>
        <td className="whitespace-nowrap px-3 py-3 text-center text-slate-700">{row.cb || "-"}</td>
      </tr>
    );
  }

  return (
    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
      <table className="min-w-full text-sm">
        <thead className="bg-slate-50/90">
          <tr>
            <th rowSpan={2} className="whitespace-nowrap px-3 py-3 text-left text-sm font-semibold text-slate-700">Sr. No.</th>
            <th rowSpan={2} className="whitespace-nowrap px-3 py-3 text-left text-sm font-semibold text-slate-700">Category</th>
            <th rowSpan={2} className="whitespace-nowrap px-3 py-3 text-center text-sm font-semibold text-slate-700">OB</th>
            <th colSpan={3} className="whitespace-nowrap px-3 py-3 text-center text-sm font-semibold text-slate-700">Addition</th>
            <th colSpan={3} className="whitespace-nowrap px-3 py-3 text-center text-sm font-semibold text-slate-700">Deduction</th>
            <th rowSpan={2} className="whitespace-nowrap px-3 py-3 text-center text-sm font-semibold text-slate-700">CB</th>
          </tr>
          <tr>
            <th className="whitespace-nowrap px-3 py-2 text-center text-xs font-semibold text-slate-700">B</th>
            <th className="whitespace-nowrap px-3 py-2 text-center text-xs font-semibold text-slate-700">P</th>
            <th className="whitespace-nowrap px-3 py-2 text-center text-xs font-semibold text-slate-700">T-in</th>
            <th className="whitespace-nowrap px-3 py-2 text-center text-xs font-semibold text-slate-700">D</th>
            <th className="whitespace-nowrap px-3 py-2 text-center text-xs font-semibold text-slate-700">S</th>
            <th className="whitespace-nowrap px-3 py-2 text-center text-xs font-semibold text-slate-700">T-out</th>
          </tr>
        </thead>
        <tbody>
          {sections.map((section) => (
            <React.Fragment key={section.title}>
              <tr className="bg-slate-50">
                <td className="px-3 py-3"></td>
                <td colSpan={9} className="px-3 py-3 font-bold text-slate-800">{section.title}</td>
              </tr>
              {section.rows.map((row, idx) => renderRow(row, idx))}
              {renderRow(section.total, section.rows.length, true)}
            </React.Fragment>
          ))}
          {renderRow(report.grandTotal, 0, true)}
        </tbody>
      </table>
    </div>
  );
}

function HerdStrengthTable({ title, rows, totalRow }) {
  const displayRows = [...rows, totalRow];
  return (
    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
      {title && <div className="border-b border-slate-100 px-4 py-3 text-sm font-semibold text-slate-700">{title}</div>}
      <table className="min-w-full text-sm">
        <thead className="bg-slate-50/90">
          <tr>
            <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Sr. No.</th>
            <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Type of Animal</th>
            <th className="whitespace-nowrap px-4 py-3 text-center text-sm font-semibold text-slate-700">Male</th>
            <th className="whitespace-nowrap px-4 py-3 text-center text-sm font-semibold text-slate-700">Female</th>
            <th className="whitespace-nowrap px-4 py-3 text-center text-sm font-semibold text-slate-700">Total</th>
          </tr>
        </thead>
        <tbody>
          {displayRows.map((row, idx) => {
            const isTotal = row.label === "Total";
            return (
              <tr key={`${row.label}-${idx}`} className={isTotal ? "bg-emerald-50/60 font-semibold" : "hover:bg-emerald-50/40"}>
                <td className="whitespace-nowrap px-4 py-3 text-slate-700">{isTotal ? "" : idx + 1}</td>
                <td className="whitespace-nowrap px-4 py-3 text-slate-700 font-medium">{row.label}</td>
                <td className="whitespace-nowrap px-4 py-3 text-center text-slate-700">{row.male || "-"}</td>
                <td className="whitespace-nowrap px-4 py-3 text-center text-slate-700">{row.female || "-"}</td>
                <td className="whitespace-nowrap px-4 py-3 text-center text-slate-700">{row.total || "-"}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function DashboardSingleTable({ rows }) {
  const rowOrder = ["Conceived", "Pending", "Inseminated but Not pregnant", "Not inseminated", "Colostrum", "Total"];
  return (
    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
      <table className="min-w-full text-sm">
        <thead className="bg-slate-50/90">
          <tr>
            <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Reproductive status</th>
            <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Count</th>
          </tr>
        </thead>
        <tbody>
          {rowOrder.map((label) => (
            <tr key={label}>
              <td className="whitespace-nowrap px-4 py-3 text-slate-700 font-medium">{label}</td>
              <td className="whitespace-nowrap px-4 py-3 text-slate-700">{rows[label] ?? 0}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function AnimalDataRecordingApp() {
  const [animals, setAnimals] = useState(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return initialAnimals;
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : initialAnimals;
    } catch {
      return initialAnimals;
    }
  });
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [selectedId, setSelectedId] = useState(null);
  const [editAnimalId, setEditAnimalId] = useState(null);
  const [detailTab, setDetailTab] = useState("pedigree");
  const [healthSubTab, setHealthSubTab] = useState("bodyWeight");
  const [maleProgenySubTab, setMaleProgenySubTab] = useState("female");
  const [daughterPerfSubTab, setDaughterPerfSubTab] = useState("production");
  const [herdView, setHerdView] = useState("current");
  const [sidebarView, setSidebarView] = useState("home");
  const [currentHerdSex, setCurrentHerdSex] = useState("Female");
  const [archiveHerdSex, setArchiveHerdSex] = useState("Female");
  const [newAnimal, setNewAnimal] = useState({ ...emptyAnimal });
  const [showSummary, setShowSummary] = useState(false);
  const [selectedFinancialYear, setSelectedFinancialYear] = useState("");
  const [selectedFinancialMonth, setSelectedFinancialMonth] = useState("Overall");
  const [selectedHerdStructureMonth, setSelectedHerdStructureMonth] = useState("April");
  const [selectedCalendarYear, setSelectedCalendarYear] = useState("");
  const [selectedCalendarMonth, setSelectedCalendarMonth] = useState("Overall");
  const [selectedConceptionDetailMonth, setSelectedConceptionDetailMonth] = useState(null);
  const [selectedBullAIYear, setSelectedBullAIYear] = useState("All");
  const [summaryDetail, setSummaryDetail] = useState(null);
  const [showMonthlyTLMBreakup, setShowMonthlyTLMBreakup] = useState(false);
  const [selectedWAFriday, setSelectedWAFriday] = useState(getNextFridayDisplay(new Date()));
  const [dailyReportDate, setDailyReportDate] = useState(formatDateDisplay(new Date()));
  const [dailyReportTab, setDailyReportTab] = useState("ai");
  const [expandedHomeAlert, setExpandedHomeAlert] = useState("");
  const [dailyReportDraftRows, setDailyReportDraftRows] = useState({});
  const [dailyReportMessage, setDailyReportMessage] = useState("");
  const [waUploadMessage, setWAUploadMessage] = useState("");
  const [waEntryRows, setWAEntryRows] = useState({});
  const [closedParityEdit, setClosedParityEdit] = useState({});
  const [expandedClosedCalving, setExpandedClosedCalving] = useState({});
  const [expandedClosedRepro, setExpandedClosedRepro] = useState({});
  const [expandedProductionParity, setExpandedProductionParity] = useState({});
  const waUploadInputRef = useRef(null);

  const normalizedAnimals = useMemo(() => animals.map((a) => withDefaults(a)), [animals]);

  useEffect(() => {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(animals));
    } catch {
      // ignore storage errors in browser-only mode
    }
  }, [animals]);

  const activeAnimals = useMemo(() => normalizedAnimals.filter((a) => !isArchivedAnimal(a)), [normalizedAnimals]);
  const archivedAnimals = useMemo(() => normalizedAnimals.filter((a) => isArchivedAnimal(a)), [normalizedAnimals]);

  const todayHeaderLabel = useMemo(() => {
    try {
      return new Date().toLocaleDateString("en-IN", { weekday: "short", day: "2-digit", month: "short", year: "numeric" });
    } catch {
      return formatDateDisplay(new Date());
    }
  }, []);

  const filteredCurrentAnimals = useMemo(() => {
    const q = search.toLowerCase();
    return activeAnimals.filter((a) => [a.tagNo, a.breed, a.category, a.status, a.isBreedingBull, a.breedingSet, a.currentSet, a.isPtBull, getFemaleLifecycle(a)].join(" ").toLowerCase().includes(q));
  }, [activeAnimals, search]);

  const filteredArchivedAnimals = useMemo(() => {
    const q = search.toLowerCase();
    return archivedAnimals.filter((a) => [a.tagNo, a.breed, a.category, a.status, a.exitDate, a.exitReason].join(" ").toLowerCase().includes(q));
  }, [archivedAnimals, search]);

  const femaleAnimals = useMemo(() => filteredCurrentAnimals.filter((a) => a.category === "Female").slice().sort(sortByTag), [filteredCurrentAnimals]);
  const maleAnimals = useMemo(() => filteredCurrentAnimals.filter((a) => a.category === "Male").slice().sort(sortByTag), [filteredCurrentAnimals]);
  const archivedFemaleAnimals = useMemo(() => filteredArchivedAnimals.filter((a) => a.category === "Female").slice().sort(sortByTag), [filteredArchivedAnimals]);
  const archivedMaleAnimals = useMemo(() => filteredArchivedAnimals.filter((a) => a.category === "Male").slice().sort(sortByTag), [filteredArchivedAnimals]);

  const registryAnimals = useMemo(() => {
    if (sidebarView === "archives") {
      return archiveHerdSex === "Female" ? archivedFemaleAnimals : archivedMaleAnimals;
    }
    return currentHerdSex === "Female" ? femaleAnimals : maleAnimals;
  }, [sidebarView, archiveHerdSex, currentHerdSex, archivedFemaleAnimals, archivedMaleAnimals, femaleAnimals, maleAnimals]);

  const registryTitle = useMemo(() => {
    if (sidebarView === "archives") return archiveHerdSex === "Female" ? "Archived Female Animals" : "Archived Male Animals";
    return currentHerdSex === "Female" ? "Current Female Animals" : "Current Male Animals";
  }, [sidebarView, archiveHerdSex, currentHerdSex]);

  const stats = useMemo(() => {
    const females = activeAnimals.filter((a) => a.category === "Female");
    const males = activeAnimals.filter((a) => a.category === "Male");
    return {
      totalAnimals: activeAnimals.length,
      femaleCount: females.length,
      maleCount: males.length,
      heiferCount: females.filter((a) => getFemaleLifecycle(a) === "Heifer").length,
      colostrumHeiferCount: females.filter((a) => getFemaleLifecycle(a) === "Colostrum-Heifer").length,
      colostrumCount: females.filter((a) => getFemaleLifecycle(a) === "Colostrum").length,
      milkCount: females.filter((a) => getFemaleLifecycle(a) === "Milk").length,
      dryCount: females.filter((a) => getFemaleLifecycle(a) === "Dry").length,
    };
  }, [activeAnimals]);

  const homeKpis = useMemo(() => {
    const females = activeAnimals.filter((a) => a.category === "Female");
    const milkingFemales = females.filter((a) => getFemaleLifecycle(a) === "Milk");

    function hasOpenPregnancy(animal) {
      const parities = animal.femaleDetails?.reproductionParities || [];
      return parities.some((parity) => {
        const conceivedRecord = (parity.aiRecords || []).find((record) => record.result === "Conceived");
        const conceptionDate = parity.conceptionDate || conceivedRecord?.aiDate || "";
        return conceptionDate && !hasCalvingAfterDate(animal, conceptionDate);
      });
    }

    const pregnantFemales = females.filter((animal) => hasOpenPregnancy(animal));

    const latestMilkTotal = milkingFemales.reduce((sum, animal) => {
      const lactations = animal.femaleDetails?.productionLactations || [];
      const selectedParity = animal.femaleDetails?.selectedProductionParity;
      const selectedLactation = lactations.find((l) => String(l.parityNo) === String(selectedParity)) || lactations[lactations.length - 1];
      const fridayRecords = selectedLactation?.fridayRecords || [];
      const latestRecord = [...fridayRecords].filter((r) => r.totalDailyYield !== "" && r.totalDailyYield !== undefined).pop();
      return sum + Number(latestRecord?.totalDailyYield || 0);
    }, 0);
    const avgMilk = milkingFemales.length ? latestMilkTotal / milkingFemales.length : 0;
    return [
      { label: "Total animals", value: activeAnimals.length, note: "Active herd", icon: "🐃" },
      { label: "In milk", value: milkingFemales.length, note: "Current lactating", icon: "🥛" },
      { label: "Pregnant", value: pregnantFemales.length, note: "Open pregnancies", icon: "✅" },
      { label: "Avg milk", value: avgMilk ? `${avgMilk.toFixed(1)} kg` : "0.0 kg", note: "Latest test-day average", icon: "📈" },
    ];
  }, [activeAnimals]);



  const homeDecisionSupport = useMemo(() => {
    const today = new Date();
    const females = activeAnimals.filter((a) => a.category === "Female");

    function daysFrom(dateStr) {
      const dt = parseDisplayDate(dateStr || "");
      if (!dt) return null;
      return Math.round((today.getTime() - dt.getTime()) / 86400000);
    }

    function daysUntil(dateStr) {
      const dt = parseDisplayDate(dateStr || "");
      if (!dt) return null;
      return Math.round((dt.getTime() - today.getTime()) / 86400000);
    }

    function getCurrentReproParity(animal) {
      const parities = animal.femaleDetails?.reproductionParities || [];
      const latestCalvingParityNo = getLatestCalvingParityNo(animal);

      // After a female has calved, the current reproductive parity must match the latest calving parity.
      // Example: after first calving, current reproductive parity is parity 1, not heifer/repro parity 0.
      if (latestCalvingParityNo > 0) {
        return parities.find((p) => Number(p.parityNo) === Number(latestCalvingParityNo)) || null;
      }

      return parities.find((p) => Number(p.parityNo) === 0) || parities[0] || null;
    }

    function getReproParityByNumber(animal, parityNo) {
      const parities = animal.femaleDetails?.reproductionParities || [];
      return parities.find((p) => Number(p.parityNo) === Number(parityNo)) || null;
    }

    function sortedAI(parity) {
      return [...(parity?.aiRecords || [])]
        .filter((record) => record.aiDate)
        .sort((a, b) => {
          const ad = parseDisplayDate(a.aiDate);
          const bd = parseDisplayDate(b.aiDate);
          if (!ad || !bd) return 0;
          return ad.getTime() - bd.getTime();
        });
    }

    function getOpenConceptionDate(animal, parity) {
      if (!parity) return "";
      const aiRecords = sortedAI(parity);
      const conceivedRecord = aiRecords.find((record) => record.result === "Conceived");
      const conceptionDate = parity.conceptionDate || conceivedRecord?.aiDate || "";
      if (!conceptionDate) return "";
      return hasCalvingAfterDate(animal, conceptionDate) ? "" : conceptionDate;
    }

    function hasOpenConception(animal, parity) {
      return Boolean(getOpenConceptionDate(animal, parity));
    }

    function getLastDryLactation(animal) {
      const lactations = [...(animal.femaleDetails?.productionLactations || [])]
        .filter((lactation) => lactation.dryDate)
        .sort((a, b) => {
          const ad = parseDisplayDate(a.dryDate || "");
          const bd = parseDisplayDate(b.dryDate || "");
          if (!ad || !bd) return 0;
          return bd.getTime() - ad.getTime();
        });
      return lactations[0] || null;
    }

    const dueForPD = [];
    const calvingWatch = [];
    const repeatBreeders = [];
    const notInseminated90 = [];
    const notPregnant120 = [];
    const longDryPeriod = [];

    females.forEach((animal) => {
      const currentParity = getCurrentReproParity(animal);
      const aiRecords = sortedAI(currentParity);
      const lastAI = aiRecords[aiRecords.length - 1];

      // 1. PD due: last AI in current reproductive parity is >30 days old and no result has been recorded
      if (lastAI?.aiDate) {
        const d = daysFrom(lastAI.aiDate);
        if (d !== null && d > 30 && lastAI.result !== "Conceived" && lastAI.result !== "Negative") {
          dueForPD.push({ tag: animal.tagNo || "—", detail: `Last AI ${lastAI.aiDate} · ${d} days ago` });
        }
      }

      // 2. Calving due: only for an open pregnancy that has not already resulted in a calving
      const conceptionDate = getOpenConceptionDate(animal, currentParity);
      if (conceptionDate) {
        const expectedDate = currentParity?.expectedCalvingDate || expectedCalving(conceptionDate);
        const d = daysUntil(expectedDate);
        if (d !== null && d < 30) {
          calvingWatch.push({ tag: animal.tagNo || "—", detail: d < 0 ? `Overdue by ${Math.abs(d)} days · ECD ${expectedDate}` : `Due in ${d} days · ECD ${expectedDate}` });
        }
      }

      // 3. Repeat breeder: current reproductive parity has >=4 inseminations and no open conception
      if (aiRecords.length >= 4 && !conceptionDate) {
        repeatBreeders.push({ tag: animal.tagNo || "—", detail: `${aiRecords.length} inseminations in current parity` });
      }

      // 4. Not inseminated after 90 days post-calving in current reproductive parity
      const lastCalving = getLastNormalCalving(animal);
      const daysPostCalving = lastCalving?.calvingDate ? daysFrom(lastCalving.calvingDate) : null;
      if (daysPostCalving !== null && daysPostCalving > 90 && aiRecords.length === 0) {
        notInseminated90.push({
          tag: animal.tagNo || "—",
          detail: `${daysPostCalving} days post-calving · no AI in current parity`,
        });
      }

      // 5. Not pregnant after 120 days post-calving in current reproductive parity
      if (daysPostCalving !== null && daysPostCalving > 120 && !hasOpenConception(animal, currentParity)) {
        notPregnant120.push({
          tag: animal.tagNo || "—",
          detail: `${daysPostCalving} days post-calving · not pregnant in current parity`,
        });
      }

      // 6. Long dry period: dried off in last parity, no later productive calving, and not pregnant yet.
      // If the animal has already calved after the dry date, the dry period has ended and must not be flagged.
      const lastDryLactation = getLastDryLactation(animal);
      if (lastDryLactation?.dryDate) {
        const dryDate = parseDisplayDate(lastDryLactation.dryDate || "");
        const dryDays = daysFrom(lastDryLactation.dryDate);
        const laterProductiveCalving = getAllCalvingOutcomeEvents(animal).some((event) => {
          const calvingDate = parseDisplayDate(event.calvingDate || "");
          return dryDate && calvingDate && calvingDate.getTime() > dryDate.getTime() && isCompleteProductiveCalvingEvent(event);
        });
        const nextReproParityNo = Number(lastDryLactation.parityNo || 0) + 1;
        const nextReproParity = getReproParityByNumber(animal, nextReproParityNo) || currentParity;
        const dryPeriodThresholdDays = 90;
        if (!laterProductiveCalving && dryDays !== null && dryDays > dryPeriodThresholdDays && !hasOpenConception(animal, nextReproParity)) {
          longDryPeriod.push({
            tag: animal.tagNo || "—",
            detail: `Dry since ${lastDryLactation.dryDate} · ${dryDays} days · not pregnant in next repro parity`,
          });
        }
      }
    });

    const alerts = [
      { id: "pd", title: "Animals due for PD", tone: "emerald", items: dueForPD },
      { id: "calving", title: "Animals due for calving", tone: "amber", items: calvingWatch },
      { id: "repeat", title: "Repeat breeder alerts", tone: "rose", items: repeatBreeders },
      { id: "notInseminated90", title: "Not inseminated >90 days post-calving", tone: "orange", items: notInseminated90 },
      { id: "notPregnant120", title: "Not pregnant >120 days post-calving", tone: "red", items: notPregnant120 },
      { id: "longDry", title: "Long dry period animals", tone: "slate", items: longDryPeriod },
    ];

    return { alerts };
  }, [activeAnimals]);

  const herdSnapshotRows = useMemo(() => buildHerdSnapshotRows(activeAnimals), [activeAnimals]);
  const milkStatusSummary = useMemo(() => buildMilkStatusSummary(activeAnimals), [activeAnimals]);
  const dryStatusSummary = useMemo(() => buildDryStatusSummary(activeAnimals), [activeAnimals]);
  const heiferStatusSummary = useMemo(() => buildHeiferStatusSummary(activeAnimals), [activeAnimals]);

  const selectedAnimal = normalizedAnimals.find((a) => a.id === selectedId) || null;
  const uiSelectedAnimal = (sidebarView === "current" || sidebarView === "archives") ? (selectedAnimal && registryAnimals.some((a) => a.id === selectedAnimal.id) ? selectedAnimal : null) : selectedAnimal;
  const activeAnimal = uiSelectedAnimal;
  const activePedigree = useMemo(() => (activeAnimal ? derivePedigreeFromKnownRelatives(activeAnimal, normalizedAnimals) : emptyPedigree), [activeAnimal, normalizedAnimals]);

  useEffect(() => {
    if (!activeAnimal) return;
    const currentPedigree = getPedigreeObjectForAnimal(activeAnimal);
    const derivedValues = {};
    Object.keys(emptyPedigree).forEach((key) => {
      if (!String(currentPedigree[key] || "").trim() && String(activePedigree[key] || "").trim()) {
        derivedValues[key] = activePedigree[key];
      }
    });
    if (!Object.keys(derivedValues).length) return;
    setAnimals((prev) => prev.map((item) => {
      if (item.id !== activeAnimal.id) return item;
      const normalized = withDefaults(item);
      if (normalized.category === "Female") {
        return { ...normalized, femaleDetails: { ...normalized.femaleDetails, pedigree: { ...normalized.femaleDetails.pedigree, ...derivedValues } } };
      }
      if (normalized.category === "Male") {
        return { ...normalized, maleDetails: { ...normalized.maleDetails, pedigree: { ...normalized.maleDetails.pedigree, ...derivedValues } } };
      }
      return normalized;
    }));
  }, [activeAnimal?.id, activePedigree]);
  const selectedReproParity = activeAnimal ? getSelectedReproParity(activeAnimal) : null;
  const selectedLactation = activeAnimal ? getSelectedLactation(activeAnimal) : null;
  const reproSummary = activeAnimal && selectedReproParity ? computeReproSummary(activeAnimal, selectedReproParity) : null;
  const selectedReproParityNo = selectedReproParity ? String(selectedReproParity.parityNo) : "";
  const selectedReproEditKey = activeAnimal ? `${activeAnimal.id}:repro:${selectedReproParityNo}` : "";
  const selectedReproParityClosed = activeAnimal && selectedReproParity ? isReproParityClosedByProductiveCalving(activeAnimal, selectedReproParity.parityNo) : false;
  const selectedReproParityLocked = Boolean(selectedReproParityClosed && !closedParityEdit[selectedReproEditKey]);
  const selectedReproParityMinimized = Boolean(selectedReproParityClosed && expandedClosedRepro[selectedReproEditKey] === false && !closedParityEdit[selectedReproEditKey]);
  const productionMetrics = selectedLactation ? computeProductionMetrics(selectedLactation) : { lactationLength: 0, totalLactationMilk: 0, standardLactationMilk: 0, peakYield: 0 };
  const selectedProductionParityKey = activeAnimal && selectedLactation ? `${activeAnimal.id}:production:${selectedLactation.parityNo}` : "";
  const monthlyTLMBreakup = selectedLactation ? getMonthlyMilkSplit(selectedLactation) : [];
  const latestProductionAutoOpenRef = useRef("");

  const waEligibleAnimals = useMemo(() => getFridayMilkEntryAnimals(normalizedAnimals, selectedWAFriday), [normalizedAnimals, selectedWAFriday]);
  const waFridaySummary = useMemo(() => buildWAFridayRow(normalizedAnimals, selectedWAFriday), [normalizedAnimals, selectedWAFriday]);
  const waHistoryRows = useMemo(() => getAllFridaySummaryRows(normalizedAnimals), [normalizedAnimals]);

  useEffect(() => {
    const targetFriday = normalizeDisplayDate(selectedWAFriday);
    const nextRows = {};
    getFridayMilkEntryAnimals(normalizedAnimals, targetFriday).forEach((animal) => {
      const lactation = getLactationForDate(animal, targetFriday) || getCurrentOrLatestLactation(animal);
      const existing = lactation?.fridayRecords?.find((r) => normalizeDisplayDate(r.date || "") === targetFriday);
      const morningMilk = existing?.morningMilk || "";
      const eveningMilk = existing?.eveningMilk || "";
      const totalDailyYield = morningMilk !== "" || eveningMilk !== "" ? (Number(morningMilk || 0) + Number(eveningMilk || 0)).toFixed(1) : (existing?.totalDailyYield || "");
      nextRows[animal.id] = { morningMilk, eveningMilk, totalDailyYield };
    });
    setWAEntryRows(nextRows);
  }, [normalizedAnimals, selectedWAFriday]);

  useEffect(() => {
    if (detailTab !== "production" || !activeAnimal || activeAnimal.category !== "Female") return;
    const lactations = [...(activeAnimal.femaleDetails?.productionLactations || [])].sort((a, b) => Number(a.parityNo || 0) - Number(b.parityNo || 0));
    const latest = lactations[lactations.length - 1];
    if (!latest) return;
    const latestKey = `${activeAnimal.id}:production:${latest.parityNo}`;
    const stamp = `${activeAnimal.id}:${latest.parityNo}:${lactations.length}`;
    if (latestProductionAutoOpenRef.current === stamp) return;
    latestProductionAutoOpenRef.current = stamp;
    if (String(activeAnimal.femaleDetails?.selectedProductionParity || "") !== String(latest.parityNo)) {
      selectProductionParity(latest.parityNo);
    }
    setExpandedProductionParity((prev) => ({ ...prev, [latestKey]: true }));
  }, [detailTab, selectedId, activeAnimal?.id, activeAnimal?.femaleDetails?.productionLactations?.length]);

  useEffect(() => {
    setShowMonthlyTLMBreakup(false);
  }, [selectedId, activeAnimal?.femaleDetails?.selectedProductionParity, selectedLactation?.entryMode]);
  const historyRows = activeAnimal?.category === "Female" ? computeHistoryRows(activeAnimal) : [];
  const afcValue = activeAnimal?.category === "Female" ? computeCalvingMetrics(activeAnimal, 1).afc ?? "" : "";
  const visibleTabs = activeAnimal?.category === "Female" ? FEMALE_TABS : activeAnimal?.category === "Male" && activeAnimal?.isBreedingBull === "Yes" ? MALE_TABS : [];
  const mateSelectionRows = useMemo(() => {
    if (!activeAnimal || activeAnimal.category !== "Female") return [];
    return getMateSelectionRowsAdvanced(activeAnimal, normalizedAnimals);
  }, [activeAnimal, normalizedAnimals]);

  const femaleProgenies = useMemo(() => {
    if (!activeAnimal || activeAnimal.category !== "Male" || activeAnimal.isBreedingBull !== "Yes") return [];
    return normalizedAnimals.filter((a) => a.category === "Female" && a.id !== activeAnimal.id && isProgenyOfBull(a, activeAnimal)).sort(sortByTag);
  }, [normalizedAnimals, activeAnimal]);

  const maleProgenies = useMemo(() => {
    if (!activeAnimal || activeAnimal.category !== "Male" || activeAnimal.isBreedingBull !== "Yes") return [];
    return normalizedAnimals.filter((a) => a.category === "Male" && a.id !== activeAnimal.id && isProgenyOfBull(a, activeAnimal)).sort(sortByTag);
  }, [normalizedAnimals, activeAnimal]);
  const maleDamRecord = useMemo(() => {
    if (!activeAnimal || activeAnimal.category !== "Male" || activeAnimal.isBreedingBull !== "Yes") return null;
    return getDamRecordRowsForMale(activeAnimal, normalizedAnimals);
  }, [activeAnimal, normalizedAnimals]);

  const allAIEvents = useMemo(() => getAllAIEvents(normalizedAnimals), [normalizedAnimals]);
  const financialYearOptions = useMemo(() => getFinancialYearOptions(normalizedAnimals), [normalizedAnimals]);
  const calendarYearOptions = useMemo(() => getCalendarYearOptions(normalizedAnimals), [normalizedAnimals]);

  React.useEffect(() => {
    if (!selectedFinancialYear && financialYearOptions.length) setSelectedFinancialYear(financialYearOptions[0].value);
  }, [selectedFinancialYear, financialYearOptions]);

  React.useEffect(() => {
    if (!selectedCalendarYear && calendarYearOptions.length) setSelectedCalendarYear(calendarYearOptions[0].value);
  }, [selectedCalendarYear, calendarYearOptions]);

  const selectedFinancialYearNum = Number(String(selectedFinancialYear || financialYearOptions[0]?.value || new Date().getFullYear()).slice(0, 4));
  const selectedCalendarYearNum = Number(String(selectedCalendarYear || calendarYearOptions[0]?.value || new Date().getFullYear()).slice(0, 4));

  const herdStructureReport = useMemo(() => buildHerdStrengthReport(normalizedAnimals, selectedFinancialYearNum, selectedHerdStructureMonth), [normalizedAnimals, selectedFinancialYearNum, selectedHerdStructureMonth]);
  const monthlyHerdMovementReport = useMemo(() => buildMonthlyHerdMovementReport(normalizedAnimals, selectedFinancialYearNum, selectedHerdStructureMonth), [normalizedAnimals, selectedFinancialYearNum, selectedHerdStructureMonth]);

  const bullAIYearOptions = useMemo(() => {
    if (!activeAnimal || activeAnimal.category !== "Male") return ["All"];

    const years = [...new Set(
      allAIEvents
        .filter((row) => matchesBullRecord(row, activeAnimal))
        .map((row) => getYearFromDisplayDate(row.aiDate))
        .filter(Boolean)
    )].sort((a, b) => a - b);

    return ["All", ...years.map(String)];
  }, [allAIEvents, activeAnimal]);


  const monthlyFinancialSummary = useMemo(() => {
    let rows = buildMonthlyFinancialSummary(normalizedAnimals, selectedFinancialYearNum);
    if (selectedFinancialMonth !== "Overall") {
      rows = rows.filter((r) => r.month === selectedFinancialMonth);
    }
    return rows;
  }, [normalizedAnimals, selectedFinancialYearNum, selectedFinancialMonth]);

  const annualFinancialSummary = useMemo(() => {
    const rows = buildMonthlyFinancialSummary(normalizedAnimals, selectedFinancialYearNum);
    const lactations = normalizedAnimals.filter((a) => a.category === "Female").flatMap((a) =>
      (a.femaleDetails?.productionLactations || [])
        .map((l) => ({ animal: a, lactation: l, metrics: computeProductionMetrics(l) }))
        .filter(({ lactation }) => {
          const d = parseDisplayDate(lactation.dryDate || "");
          return d && ((d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1) === selectedFinancialYearNum);
        })
    );
    const reproRows = normalizedAnimals.filter((a) => a.category === "Female").flatMap((a) => computeHistoryRows(a)).filter((r) => {
      const d = parseDisplayDate(r.dateCalved || "");
      return d && ((d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1) === selectedFinancialYearNum);
    });
    const avg = (arr) => arr.length ? (arr.reduce((s,v)=>s+Number(v||0),0)/arr.length).toFixed(2) : "0.00";
    return {
      calvings: rows.reduce((s,r)=>s+r.calvings,0),
      stillbirths: rows.reduce((s,r)=>s+r.stillbirths,0),
      abortions: rows.reduce((s,r)=>s+r.abortions,0),
      mortality: rows.reduce((s,r)=>s+r.mortality,0),
      lactationsCompleted: rows.reduce((s,r)=>s+r.lactationsCompleted,0),
      avgTLMY: avg(lactations.map((x)=>x.metrics.totalLactationMilk.toFixed ? x.metrics.totalLactationMilk.toFixed(2) : x.metrics.totalLactationMilk)),
      avgSLMY: avg(lactations.map((x)=>x.metrics.standardLactationMilk.toFixed ? x.metrics.standardLactationMilk.toFixed(2) : x.metrics.standardLactationMilk)),
      avgLL: avg(lactations.map((x)=>x.metrics.lactationLength)),
      avgPY: avg(lactations.map((x)=>x.metrics.peakYield.toFixed ? x.metrics.peakYield.toFixed(2) : x.metrics.peakYield)),
      avgAFC: avg(reproRows.map((r)=>r.afc)),
      avgSP: avg(reproRows.map((r)=>r.sp)),
      avgDP: avg(reproRows.map((r)=>r.dp)),
      avgCI: avg(reproRows.map((r)=>r.ci)),
    };
  }, [normalizedAnimals, selectedFinancialYearNum]);

  const monthlyConceptionRows = useMemo(() => {
    let rows = buildMonthlyConceptionSummary(allAIEvents, selectedCalendarYearNum);
    if (selectedCalendarMonth !== "Overall") {
      rows = rows.filter((r) => r.month === selectedCalendarMonth);
    }
    return rows;
  }, [allAIEvents, selectedCalendarYearNum, selectedCalendarMonth]);

  const annualConceptionSummary = useMemo(() => {
    const totalAI = allAIEvents.filter((r) => getYearFromDisplayDate(r.aiDate) === selectedCalendarYearNum).length;
    const conceived = allAIEvents.filter((r) => r.result === "Conceived" && getYearFromDisplayDate(r.conceptionDate || r.aiDate) === selectedCalendarYearNum).length;
    return { totalAI, conceived, crPct: totalAI ? ((conceived / totalAI) * 100).toFixed(2) : "0.00" };
  }, [allAIEvents, selectedCalendarYearNum]);

  const conceptionDetailRows = useMemo(() => {
    if (!selectedConceptionDetailMonth) return [];
    return buildMonthlyConceptionDetailRows(normalizedAnimals, allAIEvents, selectedCalendarYearNum, selectedConceptionDetailMonth);
  }, [normalizedAnimals, allAIEvents, selectedCalendarYearNum, selectedConceptionDetailMonth]);

  const quarterWiseConceptionRows = useMemo(() => buildQuarterWiseConceptionSummary(allAIEvents, selectedCalendarYearNum), [allAIEvents, selectedCalendarYearNum]);
  const bullWiseConceptionRows = useMemo(() => buildBullWiseConceptionSummary(allAIEvents, selectedCalendarYearNum), [allAIEvents, selectedCalendarYearNum]);
  const femaleConceptionByNthAI = useMemo(() => buildFemaleConceptionByNthAI(allAIEvents, selectedCalendarYearNum), [allAIEvents, selectedCalendarYearNum]);

  function exportAnalyticsSummaryPdf() {
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();
    const marginX = 10;
    const pageHeight = doc.internal.pageSize.getHeight();

    const addTitleBlock = () => {
      doc.setFontSize(15);
      doc.text("ANALYTICAL SUMMARY", pageWidth / 2, 12, { align: "center" });
      doc.setFontSize(9);
      doc.text(`Financial Year: ${financialYearOptions.find((o) => String(o.value) === String(selectedFinancialYear))?.label || String(selectedFinancialYear || "")}`, marginX, 20);
      doc.text(`Financial Month: ${selectedFinancialMonth || "Overall"}`, 90, 20);
      doc.text(`Calendar Year: ${selectedCalendarYearNum || ""}`, 170, 20);
      doc.text(`Calendar Month: ${selectedCalendarMonth || "Overall"}`, 235, 20);
    };

    const ensureSpace = (needed = 18) => {
      const currentY = (doc.lastAutoTable?.finalY || 24) + 6;
      if (currentY + needed > pageHeight - 12) {
        doc.addPage();
        return 14;
      }
      return currentY;
    };

    const addSection = (title, head, body, options = {}) => {
      const startY = ensureSpace(options.estimatedHeight || 22);
      doc.setFontSize(11);
      doc.text(title, marginX, startY);
      autoTable(doc, {
        startY: startY + 3,
        head: [head],
        body,
        theme: "grid",
        styles: { fontSize: options.fontSize || 8, cellPadding: 1.6, overflow: "linebreak" },
        headStyles: { fillColor: [220, 245, 232], textColor: 20, fontStyle: "bold" },
        margin: { left: marginX, right: marginX },
        ...options.tableOptions,
      });
    };

    addTitleBlock();

    addSection(
      "Monthly Summary · Financial Year",
      ["Month", "Calvings", "Stillbirths", "Abortions", "Mortality", "Lactations Completed"],
      monthlyFinancialSummary.map((r) => [r.month, r.calvings, r.stillbirths, r.abortions, r.mortality, r.lactationsCompleted]),
      { estimatedHeight: 55 }
    );

    if (summaryDetail) {
      const summaryRows = buildMonthlyEventDetails(normalizedAnimals, selectedFinancialYearNum, summaryDetail.month, summaryDetail.type);
      const detailHead = summaryDetail.type === "mortality"
        ? ["Animal No.", "Event Date", "Status", "Category", "Reason"]
        : summaryDetail.type === "lactationsCompleted"
        ? ["Animal No.", "Dry Date", "Parity", "TLMY", "SLMY", "LL", "PY"]
        : ["Animal No.", "Event Date", "Parity", "Sex of Calf", "Calf Tag"];
      const detailBody = summaryRows.map((r) => summaryDetail.type === "mortality"
        ? [r.tagNo || "—", r.eventDate || "—", r.status || "—", r.category || "—", r.reason || "—"]
        : summaryDetail.type === "lactationsCompleted"
        ? [r.tagNo || "—", r.dryDate || "—", r.parity || "—", r.tlmy || "—", r.slmy || "—", r.ll || "—", r.py || "—"]
        : [r.tagNo || "—", r.eventDate || "—", r.parity || "—", r.sexOfCalf || "—", r.calfTag || "—"]);
      addSection(
        `Financial Year Detail · ${summaryDetail.month} · ${summaryDetail.type}`,
        detailHead,
        detailBody.length ? detailBody : [["No records", "", "", "", ""]],
        { estimatedHeight: 40, fontSize: 7.5 }
      );
    }

    addSection(
      "Conception Dashboard · Calendar Year",
      ["Month", "Total AI", "Total Conceived", "Conception Rate %"],
      monthlyConceptionRows.map((r) => [r.month, r.totalAI, r.conceived, r.crPct]),
      { estimatedHeight: 55 }
    );

    addSection(
      "Conception Totals",
      ["Total AI", "Conceived", "CR %"],
      [[annualConceptionSummary.totalAI, annualConceptionSummary.conceived, annualConceptionSummary.crPct]],
      { estimatedHeight: 20 }
    );

    if (selectedConceptionDetailMonth) {
      addSection(
        `AI Details · ${selectedConceptionDetailMonth} ${selectedCalendarYearNum}`,
        ["Animal inseminated", "Date of last calving", "No. of AI so far", "Bull used for current AI", "Date of current AI", "Result of insemination"],
        conceptionDetailRows.length
          ? conceptionDetailRows.map((row) => [row.tagNo, row.lastCalvingDate, row.totalAISoFar, row.bullUsed, row.currentAIDate, row.inseminationResult])
          : [["No AI records found for this month.", "", "", "", "", ""]],
        { estimatedHeight: 45, fontSize: 7.5 }
      );
    }

    addSection(
      "Quarter-wise Conception Rate",
      ["Quarter", "No. of AI", "Pregnant animals", "CR %"],
      quarterWiseConceptionRows.map((r) => [r.quarter, r.totalAI, r.conceived, r.crPct]),
      { estimatedHeight: 35 }
    );

    addSection(
      "Bull-wise Conception Rate",
      ["Bull No.", "Set No.", "Total No. of AI", "Total Conceived", "CR %"],
      bullWiseConceptionRows.map((r) => [r.bullNo, r.setNo || "—", r.totalAI, r.conceived, r.crPct]),
      { estimatedHeight: 45 }
    );

    addSection(
      "Female Conception Rate by AI Number (nth)",
      ["Group", "1st AI", "2nd AI", "3rd AI", "4th & above"],
      ["Heifer", "Milk+ dry"].map((grp) => [
        grp,
        `${femaleConceptionByNthAI[grp][1].ai}/${femaleConceptionByNthAI[grp][1].conceived} (${femaleConceptionByNthAI[grp][1].crPct}%)`,
        `${femaleConceptionByNthAI[grp][2].ai}/${femaleConceptionByNthAI[grp][2].conceived} (${femaleConceptionByNthAI[grp][2].crPct}%)`,
        `${femaleConceptionByNthAI[grp][3].ai}/${femaleConceptionByNthAI[grp][3].conceived} (${femaleConceptionByNthAI[grp][3].crPct}%)`,
        `${femaleConceptionByNthAI[grp][4].ai}/${femaleConceptionByNthAI[grp][4].conceived} (${femaleConceptionByNthAI[grp][4].crPct}%)`,
      ]),
      { estimatedHeight: 28 }
    );

    doc.save(`analytical-summary-${selectedCalendarYearNum || "report"}.pdf`);
  }

  const selectedBullAIEvents = useMemo(() => {
    if (!activeAnimal) return [];
    if (activeAnimal.category !== "Male") return [];
    if (!allAIEvents || !Array.isArray(allAIEvents)) return [];

    let rows = allAIEvents.filter((row) => matchesBullRecord(row, activeAnimal));

    if (selectedBullAIYear !== "All") {
      rows = rows.filter((row) => getYearFromDisplayDate(row.aiDate) === Number(selectedBullAIYear));
    }

    return rows.sort((a, b) => {
      const ad = parseDisplayDate(a.aiDate);
      const bd = parseDisplayDate(b.aiDate);
      if (!ad || !bd) return 0;
      return bd.getTime() - ad.getTime();
    });
  }, [allAIEvents, selectedAnimal, selectedBullAIYear]);

  const dailyReportTabs = [
    { id: "ai", label: "AI information" },
    { id: "calvings", label: "Calvings" },
    { id: "abortion", label: "Abortion" },
    { id: "stillbirth", label: "Stillbirth" },
    { id: "pd", label: "Pregnancy Diagnosis" },
    { id: "treatment", label: "Treatment" },
  ];

  const dailyReportRows = useMemo(() => {
    const selectedDate = normalizeDisplayDate(dailyReportDate || "");
    const femaleAnimals = normalizedAnimals.filter((a) => a.category === "Female");
    const aiRows = (allAIEvents || []).filter((row) => normalizeDisplayDate(row.aiDate || "") === selectedDate);
    const calvingEvents = femaleAnimals.flatMap((animal) =>
      (animal.femaleDetails?.calvingParities || [])
        .filter((row) => normalizeDisplayDate(row.calvingDate || "") === selectedDate)
        .map((row) => ({ animal, row }))
    );
    const treatmentRows = normalizedAnimals.flatMap((animal) => {
      const health = animal.category === "Female" ? animal.femaleDetails?.health : animal.maleDetails?.health;
      return (health?.treatmentRecords || [])
        .filter((row) => normalizeDisplayDate(row.treatmentDate || "") === selectedDate)
        .map((row) => ({ animalNo: animal.tagNo || "—", diagnosis: row.diagnosis || "—", treatmentGiven: row.treatmentGiven || "—" }));
    });
    return {
      ai: aiRows.map((row) => ({ animalNo: row.tagNo || "—", bullNo: row.bullNo || row.aiBullNo || "—", setNo: row.setNo || row.aiSetNo || "—" })),
      calvings: calvingEvents.filter(({ row }) => (row.calvingOutcome || "Normal calving") === "Normal calving").map(({ animal, row }) => ({ animalNo: animal.tagNo || "—", calfSex: row.calfSex || "—", calfTagNo: row.calfTag || "—" })),
      abortion: calvingEvents.filter(({ row }) => row.calvingOutcome === "Abortion").map(({ animal }) => ({ animalNo: animal.tagNo || "—" })),
      stillbirth: calvingEvents.filter(({ row }) => row.calvingOutcome === "Stillbirth").map(({ animal }) => ({ animalNo: animal.tagNo || "—" })),
      pd: aiRows.filter((row) => ["Conceived", "Negative"].includes(row.result || "")).map((row) => ({ animalNo: row.tagNo || "—", inseminationDate: row.aiDate || "—", result: row.result === "Conceived" ? "Positive" : "Negative" })),
      treatment: treatmentRows,
    };
  }, [normalizedAnimals, allAIEvents, dailyReportDate]);

  const dailyReportTable = useMemo(() => {
    const tables = {
      ai: { headers: ["Animal No.", "Bull No.", "Set"], keys: ["animalNo", "bullNo", "setNo"] },
      calvings: { headers: ["Animal No.", "Calf sex", "Calf Tag No."], keys: ["animalNo", "calfSex", "calfTagNo"] },
      abortion: { headers: ["Animal No."], keys: ["animalNo"] },
      stillbirth: { headers: ["Animal No."], keys: ["animalNo"] },
      pd: { headers: ["Animal No.", "Date of insemination", "Result"], keys: ["animalNo", "inseminationDate", "result"] },
      treatment: { headers: ["Animal No.", "Diagnosis", "Treatment given"], keys: ["animalNo", "diagnosis", "treatmentGiven"] },
    };
    return tables[dailyReportTab] || tables.ai;
  }, [dailyReportTab]);



  const dailyReportBlankRows = {
    ai: { animalNo: "", bullNo: "", setNo: "" },
    calvings: { animalNo: "", calfSex: "", calfTagNo: "" },
    abortion: { animalNo: "" },
    stillbirth: { animalNo: "" },
    pd: { animalNo: "", inseminationDate: "", result: "Positive" },
    treatment: { animalNo: "", diagnosis: "", treatmentGiven: "" },
  };

  function getDailyReportDraftKey(tab = dailyReportTab, date = dailyReportDate) {
    return `${normalizeDisplayDate(date || "")}__${tab}`;
  }

  function getDailyReportEditableRows() {
    const key = getDailyReportDraftKey();
    if (dailyReportDraftRows[key]) return dailyReportDraftRows[key];
    const sourceRows = dailyReportRows[dailyReportTab] || [];
    return sourceRows.length ? sourceRows.map((row) => ({ ...row })) : [{ ...(dailyReportBlankRows[dailyReportTab] || {}) }];
  }

  function setDailyReportEditableRows(nextRows) {
    const key = getDailyReportDraftKey();
    setDailyReportDraftRows((prev) => ({ ...prev, [key]: nextRows }));
    setDailyReportMessage("");
  }

  function updateDailyReportCell(rowIdx, field, value) {
    const rows = getDailyReportEditableRows();
    const nextRows = rows.map((row, idx) => idx === rowIdx ? { ...row, [field]: value } : row);
    setDailyReportEditableRows(nextRows);
  }

  function addDailyReportRow() {
    const rows = getDailyReportEditableRows();
    setDailyReportEditableRows([...(rows || []), { ...(dailyReportBlankRows[dailyReportTab] || {}) }]);
  }

  function removeDailyReportRow(rowIdx) {
    const rows = getDailyReportEditableRows();
    const nextRows = rows.filter((_, idx) => idx !== rowIdx);
    setDailyReportEditableRows(nextRows.length ? nextRows : [{ ...(dailyReportBlankRows[dailyReportTab] || {}) }]);
  }

  function saveDailyReportRows() {
    const selectedDate = normalizeDisplayDate(dailyReportDate || "");
    if (!selectedDate) {
      setDailyReportMessage("Please select a valid date before saving.");
      return;
    }

    const rows = getDailyReportEditableRows()
      .map((row) => ({ ...row, animalNo: String(row.animalNo || "").trim() }))
      .filter((row) => row.animalNo);

    if (!rows.length) {
      setDailyReportMessage("Please enter at least one animal number before saving.");
      return;
    }

    let saved = 0;
    const missing = [];

    setAnimals((prev) => {
      let mapped = prev.map((animal) => withDefaults(animal));

      rows.forEach((row) => {
        const animalIdx = mapped.findIndex((animal) => String(animal.tagNo || "").trim().toLowerCase() === String(row.animalNo || "").trim().toLowerCase());
        if (animalIdx < 0) {
          missing.push(row.animalNo);
          return;
        }

        const animal = withDefaults(mapped[animalIdx]);

        if (dailyReportTab === "ai") {
          if (animal.category !== "Female") {
            missing.push(`${row.animalNo} (not female)`);
            return;
          }
          const parities = [...(animal.femaleDetails?.reproductionParities || [makeReproParity(0)])];
          let idx = parities.findIndex((p) => String(p.parityNo) === String(animal.femaleDetails?.selectedReproParity ?? ""));
          if (idx < 0) idx = parities.length - 1;
          const aiRecord = {
            aiDate: selectedDate,
            aiBullNo: String(row.bullNo || "").trim(),
            aiSetNo: String(row.setNo || "").trim(),
            result: "Pending",
          };
          parities[idx] = syncReproConceptionFields({ ...parities[idx], aiRecords: [...(parities[idx].aiRecords || []), aiRecord] });
          mapped[animalIdx] = withDefaults({ ...animal, femaleDetails: { ...animal.femaleDetails, reproductionParities: parities } });
          saved += 1;
          return;
        }

        if (dailyReportTab === "calvings" || dailyReportTab === "abortion" || dailyReportTab === "stillbirth") {
          if (animal.category !== "Female") {
            missing.push(`${row.animalNo} (not female)`);
            return;
          }
          const parities = [...(animal.femaleDetails?.calvingParities || [])];
          const outcome = dailyReportTab === "abortion" ? "Abortion" : dailyReportTab === "stillbirth" ? "Stillbirth" : "Normal calving";
          const nextParityNo = String((parities.length || 0) + 1);
          const calvingRow = {
            ...makeCalvingParity(nextParityNo),
            calvingDate: selectedDate,
            calvingOutcome: outcome,
            calfSex: outcome === "Normal calving" ? (row.calfSex || "") : "",
            calfTag: outcome === "Normal calving" ? (row.calfTagNo || "") : "",
            calfSire: "",
          };
          mapped[animalIdx] = withDefaults({ ...animal, femaleDetails: { ...animal.femaleDetails, calvingParities: [...parities, calvingRow] } });
          saved += 1;
          return;
        }

        if (dailyReportTab === "pd") {
          if (animal.category !== "Female") {
            missing.push(`${row.animalNo} (not female)`);
            return;
          }
          const pdResult = String(row.result || "").toLowerCase().startsWith("p") ? "Conceived" : "Negative";
          const targetAIDate = normalizeDisplayDate(row.inseminationDate || "");
          const parities = [...(animal.femaleDetails?.reproductionParities || [makeReproParity(0)])];
          let updated = false;
          for (let p = 0; p < parities.length; p += 1) {
            const aiRecords = [...(parities[p].aiRecords || [])];
            const aiIdx = aiRecords.findIndex((r) => normalizeDisplayDate(r.aiDate || "") === targetAIDate);
            if (aiIdx >= 0) {
              aiRecords[aiIdx] = { ...aiRecords[aiIdx], result: pdResult };
              parities[p] = syncReproConceptionFields({ ...parities[p], aiRecords });
              updated = true;
              break;
            }
          }
          if (!updated) {
            let idx = parities.findIndex((p) => String(p.parityNo) === String(animal.femaleDetails?.selectedReproParity ?? ""));
            if (idx < 0) idx = parities.length - 1;
            parities[idx] = syncReproConceptionFields({
              ...parities[idx],
              aiRecords: [...(parities[idx].aiRecords || []), { aiDate: targetAIDate || selectedDate, aiBullNo: "", aiSetNo: "", result: pdResult }],
            });
          }
          mapped[animalIdx] = withDefaults({ ...animal, femaleDetails: { ...animal.femaleDetails, reproductionParities: parities } });
          saved += 1;
          return;
        }

        if (dailyReportTab === "treatment") {
          const isFemale = animal.category === "Female";
          const health = isFemale ? animal.femaleDetails?.health : animal.maleDetails?.health;
          const nextHealth = {
            ...health,
            treatmentRecords: [
              ...(health?.treatmentRecords || []),
              {
                treatmentDate: selectedDate,
                diagnosis: row.diagnosis || "",
                treatmentGiven: row.treatmentGiven || "",
              },
            ],
          };
          mapped[animalIdx] = isFemale
            ? withDefaults({ ...animal, femaleDetails: { ...animal.femaleDetails, health: nextHealth } })
            : withDefaults({ ...animal, maleDetails: { ...animal.maleDetails, health: nextHealth } });
          saved += 1;
        }
      });

      return mapped;
    });

    const key = getDailyReportDraftKey();
    setDailyReportDraftRows((prev) => ({ ...prev, [key]: getDailyReportEditableRows() }));
    setDailyReportMessage(`Saved ${saved} row(s).${missing.length ? ` Not found/skipped: ${missing.join(", ")}.` : ""}`);
  }


  function updateWAMilkEntry(animalId, field, value) {
    const cleanValue = parseMilkEntryValue(value);
    setWAEntryRows((prev) => {
      const current = prev[animalId] || { morningMilk: "", eveningMilk: "", totalDailyYield: "" };
      const next = { ...current, [field]: cleanValue };
      next.totalDailyYield = next.morningMilk !== "" || next.eveningMilk !== "" ? (Number(next.morningMilk || 0) + Number(next.eveningMilk || 0)).toFixed(1) : "";
      return { ...prev, [animalId]: next };
    });
  }

  function saveWAInlineEntries() {
    if (!isFridayDisplayDate(selectedWAFriday)) {
      setWAUploadMessage("Please select a valid Friday date before saving milk records.");
      return;
    }
    const targetFriday = normalizeDisplayDate(selectedWAFriday);
    let savedCount = 0;
    let skippedCount = 0;

    setAnimals((prev) => {
      const baseAnimals = prev.map((animal) => withDefaults(animal));
      return baseAnimals.map((animal) => {
        if (animal.category !== "Female") return animal;
        const draft = waEntryRows[animal.id];
        if (!draft || (draft.morningMilk === "" && draft.eveningMilk === "")) return animal;
        const lactation = getLactationForDate(animal, targetFriday);
        if (!lactation) { skippedCount += 1; return animal; }
        const lactations = [...animal.femaleDetails.productionLactations];
        const lactationIdx = lactations.findIndex((l) => String(l.parityNo) === String(lactation.parityNo));
        if (lactationIdx < 0) { skippedCount += 1; return animal; }
        const existingRows = [...(lactations[lactationIdx].fridayRecords || [])];
        const rowIdx = existingRows.findIndex((r) => normalizeDisplayDate(r.date || "") === targetFriday);
        const nextRecord = recalcFridayRecord({
          ...makeFridayRecord(targetFriday),
          ...(rowIdx >= 0 ? existingRows[rowIdx] : {}),
          date: targetFriday,
          morningMilk: draft.morningMilk,
          eveningMilk: draft.eveningMilk,
        });
        if (rowIdx >= 0) existingRows[rowIdx] = nextRecord;
        else existingRows.push(nextRecord);
        existingRows.sort((a, b) => {
          const ad = parseDisplayDate(a.date || "");
          const bd = parseDisplayDate(b.date || "");
          if (!ad || !bd) return 0;
          return ad.getTime() - bd.getTime();
        });
        lactations[lactationIdx] = { ...lactations[lactationIdx], entryMode: "Friday Records", fridayRecords: existingRows };
        savedCount += 1;
        return { ...animal, femaleDetails: { ...animal.femaleDetails, productionLactations: lactations } };
      });
    });

    const parts = [`Saved ${savedCount} Friday milk record(s) for ${targetFriday}. WA/HA and animal production tabs refresh automatically.`];
    if (skippedCount) parts.push(`Skipped ${skippedCount} animal(s) without matching lactation/parity on that date.`);
    setWAUploadMessage(parts.join(" "));
  }

  function downloadWAFridayTemplate() {
    if (!isFridayDisplayDate(selectedWAFriday)) {
      setWAUploadMessage("Please select a valid Friday date before downloading the milk sheet.");
      return;
    }
    const rows = getFridayMilkEntryAnimals(normalizedAnimals, selectedWAFriday);
    const headers = ["Friday Date", "Tag No.", "Parity", "Date of Calving", "Morning Milk", "Evening Milk", "Total Milk", "Remarks"];
    const csvRows = [
      headers.join(","),
      ...rows.map((animal) => {
        const lactation = getLactationForDate(animal, selectedWAFriday) || getCurrentOrLatestLactation(animal);
        const existing = lactation?.fridayRecords?.find((r) => normalizeDisplayDate(r.date || "") === normalizeDisplayDate(selectedWAFriday));
        const cells = [
          normalizeDisplayDate(selectedWAFriday),
          animal.tagNo || "",
          lactation?.parityNo || "",
          lactation?.calvingDate || "",
          existing?.morningMilk || "",
          existing?.eveningMilk || "",
          existing?.totalDailyYield || "",
          "",
        ];
        return cells.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(",");
      }),
    ].join("\n");
    const blob = new Blob([csvRows], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `wa-ha-friday-sheet-${normalizeDisplayDate(selectedWAFriday).replaceAll("/", "-")}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    setWAUploadMessage(`Friday milk sheet downloaded for ${normalizeDisplayDate(selectedWAFriday)}.`);
  }

  function triggerWAUpload() {
    if (!isFridayDisplayDate(selectedWAFriday)) {
      setWAUploadMessage("Please select a valid Friday date before uploading a filled sheet.");
      return;
    }
    waUploadInputRef.current?.click();
  }

  function handleWAUploadFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsedRows = parseSpreadsheetFile(file, reader.result);
        if (!parsedRows.length) {
          setWAUploadMessage("The uploaded file is empty or could not be read. Please upload the generated Friday sheet.");
          event.target.value = "";
          return;
        }
        const targetFriday = normalizeDisplayDate(selectedWAFriday);
        const validRows = parsedRows.filter((row) => normalizeDisplayDate(row["Friday Date"] || row["Friday date"] || row["Date"] || "") === targetFriday && String(row["Tag No."] || row["Tag No"] || row["Animal No."] || "").trim());
        if (!validRows.length) {
          setWAUploadMessage(`No usable rows found for ${targetFriday}. Please upload the same Friday sheet after filling it.`);
          event.target.value = "";
          return;
        }

        const uploadResult = { updatedCount: 0, missingTags: [], notFound: [] };

        setAnimals((prev) => {
          const baseAnimals = prev.map((animal) => withDefaults(animal));
          const indexByTag = new Map(baseAnimals.map((animal, idx) => [String(animal.tagNo || "").trim(), idx]));
          const touchedTags = new Set();
          const nextAnimals = [...baseAnimals];

          validRows.forEach((row) => {
            const tag = String(row["Tag No."] || row["Tag No"] || row["Animal No."] || "").trim();
            if (!tag || touchedTags.has(tag)) return;
            touchedTags.add(tag);
            const idx = indexByTag.get(tag);
            if (idx === undefined) {
              uploadResult.notFound.push(tag);
              return;
            }
            const animal = nextAnimals[idx];
            if (animal.category !== "Female") return;
            const morningMilk = parseNumericValue(row["Morning Milk"] || row["Morning"] || "");
            const eveningMilk = parseNumericValue(row["Evening Milk"] || row["Evening"] || "");
            if (morningMilk === "" && eveningMilk === "") return;

            const lactation = getLactationForDate(animal, targetFriday);
            if (!lactation) {
              uploadResult.missingTags.push(tag);
              return;
            }

            const lactations = [...animal.femaleDetails.productionLactations];
            const lactationIdx = lactations.findIndex((l) => String(l.parityNo) === String(lactation.parityNo));
            if (lactationIdx < 0) {
              uploadResult.missingTags.push(tag);
              return;
            }

            const existingRows = [...(lactations[lactationIdx].fridayRecords || [])];
            const rowIdx = existingRows.findIndex((r) => normalizeDisplayDate(r.date || "") === targetFriday);
            const nextRecord = recalcFridayRecord({
              ...makeFridayRecord(targetFriday),
              ...(rowIdx >= 0 ? existingRows[rowIdx] : {}),
              date: targetFriday,
              morningMilk,
              eveningMilk,
            });
            if (rowIdx >= 0) existingRows[rowIdx] = nextRecord;
            else existingRows.push(nextRecord);
            existingRows.sort((a, b) => {
              const ad = parseDisplayDate(a.date || "");
              const bd = parseDisplayDate(b.date || "");
              if (!ad || !bd) return 0;
              return ad.getTime() - bd.getTime();
            });

            lactations[lactationIdx] = {
              ...lactations[lactationIdx],
              entryMode: "Friday Records",
              fridayRecords: existingRows,
            };

            nextAnimals[idx] = {
              ...animal,
              femaleDetails: {
                ...animal.femaleDetails,
                productionLactations: lactations,
              },
            };
            uploadResult.updatedCount += 1;
          });

          return nextAnimals;
        });

        const parts = [`Uploaded ${uploadResult.updatedCount} animal record(s) for ${targetFriday}. Herd analytics has been refreshed in the frontend.`];
        if (uploadResult.missingTags.length) parts.push(`Skipped ${uploadResult.missingTags.length} tag(s) not in milk on that Friday.`);
        if (uploadResult.notFound.length) parts.push(`Unknown tags: ${uploadResult.notFound.join(", ")}.`);
        setWAUploadMessage(parts.join(" "));
      } catch (error) {
        setWAUploadMessage("The uploaded file could not be processed. Please upload the generated CSV/XLSX Friday sheet without changing the column headings.");
      }
      event.target.value = "";
    };

    const name = String(file.name || "").toLowerCase();
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) reader.readAsArrayBuffer(file);
    else reader.readAsText(file);
  }
  const selectedBullAIStats = useMemo(() => getBullAIStats(selectedBullAIEvents), [selectedBullAIEvents]);

  useEffect(() => {
    setSelectedBullAIYear("All");
  }, [selectedAnimal?.id]);

  useEffect(() => {
    setSelectedConceptionDetailMonth(selectedCalendarMonth !== "Overall" ? selectedCalendarMonth : null);
  }, [selectedCalendarMonth, selectedCalendarYear]);

  useEffect(() => {
    if (sidebarView === "add") {
      setShowAdd(true);
      setShowSummary(false);
    } else if (sidebarView === "analytics") {
      setShowSummary(true);
      setShowAdd(false);
    } else {
      setShowSummary(false);
      if (sidebarView !== "add") setShowAdd(false);
    }
    if (sidebarView === "current") setHerdView("current");
    if (sidebarView === "archives") setHerdView("archive");
  }, [sidebarView]);

  useEffect(() => {
    if ((sidebarView === "current" || sidebarView === "archives") && selectedId !== null) {
      const existsInVisibleRegistry = registryAnimals.some((animal) => animal.id === selectedId);
      if (!existsInVisibleRegistry) {
        setSelectedId(null);
        setDetailTab("pedigree");
      }
    }
  }, [sidebarView, currentHerdSex, archiveHerdSex, registryAnimals, selectedId]);



  function handleFormStatusChange(status) {
    setNewAnimal((s) => normalizeAnimalFormData({ ...s, status }));
  }

  function handleFormCategoryChange(category) {
    setNewAnimal((s) => normalizeAnimalFormData({ ...s, category }));
  }



  function exportDairyScopeReport(reportType) {
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const today = formatDateDisplay(new Date());

    function header(title, subtitle = "") {
      doc.setFontSize(15);
      doc.text(title, 148, 13, { align: "center" });
      doc.setFontSize(9);
      doc.text("ICAR-CENTRAL INSTITUTE FOR RESEARCH ON BUFFALOES", 148, 19, { align: "center" });
      doc.text("SUB-CAMPUS, NABHA PUNJAB 147201", 148, 24, { align: "center" });
      doc.setFontSize(8);
      doc.text(`Generated: ${today}`, 14, 30);
      if (subtitle) doc.text(subtitle, 14, 35);
    }

    if (reportType === "femaleHistory") {
      if (!activeAnimal || activeAnimal.category !== "Female") {
        window.alert("Please select a female animal first, then open Reports and export Female History Sheet.");
        return;
      }
      exportHistoryPdf(activeAnimal);
      return;
    }

    if (reportType === "bullHistory") {
      if (!activeAnimal || activeAnimal.category !== "Male" || activeAnimal.isBreedingBull !== "Yes") {
        window.alert("Please select a breeding bull first, then open Reports and export Breeding Bull History Sheet.");
        return;
      }
      exportBreedingBullHistoryPdf(activeAnimal);
      return;
    }

    if (reportType === "herdStructure") {
      header("HERD STRUCTURE REPORT", `Total active animals: ${activeAnimals.length}`);
      const rows = [
        ["Total active herd", activeAnimals.length],
        ["Total females", activeAnimals.filter((a) => a.category === "Female").length],
        ["Buffaloes in milk", activeAnimals.filter((a) => a.category === "Female" && getFemaleLifecycle(a) === "Milk").length],
        ["Dry buffaloes", activeAnimals.filter((a) => a.category === "Female" && getFemaleLifecycle(a) === "Dry").length],
        ["Heifers", activeAnimals.filter((a) => a.category === "Female" && getFemaleLifecycle(a) === "Heifer").length],
        ["Total males", activeAnimals.filter((a) => a.category === "Male").length],
        ["Breeding bulls", activeAnimals.filter((a) => a.category === "Male" && a.isBreedingBull === "Yes").length],
        ["Archived animals", archivedAnimals.length],
      ];
      autoTable(doc, { startY: 42, head: [["Category", "Number"]], body: rows, theme: "grid", headStyles: { fillColor: [16, 120, 83] } });
      doc.save(`DairyScope-herd-structure-${new Date().toISOString().slice(0, 10)}.pdf`);
      return;
    }

    if (reportType === "aiFertility") {
      header("AI & FERTILITY REPORT", `Total AI records: ${allAIEvents.length}`);
      const rows = allAIEvents.map((row) => [
        row.tagNo || "",
        row.aiDate || "",
        row.bullNo || row.aiBullNo || "",
        row.setNo || row.aiSetNo || "",
        row.result || "Pending",
      ]);
      autoTable(doc, {
        startY: 42,
        head: [["Animal No.", "AI Date", "Bull No.", "Set", "Result"]],
        body: rows.length ? rows : [["No records", "", "", "", ""]],
        theme: "grid",
        headStyles: { fillColor: [16, 120, 83] },
      });
      doc.save(`DairyScope-ai-fertility-${new Date().toISOString().slice(0, 10)}.pdf`);
      return;
    }

    if (reportType === "production") {
      header("PRODUCTION REPORT", "Production lactation summary");
      const rows = normalizedAnimals
        .filter((a) => a.category === "Female")
        .flatMap((animal) => (animal.femaleDetails?.productionLactations || []).map((lactation) => {
          const metrics = computeProductionMetrics(lactation);
          return [
            animal.tagNo || "",
            lactation.parityNo || "",
            lactation.entryMode || "",
            lactation.calvingDate || "",
            lactation.dryDate || "",
            metrics.totalLactationMilk || "",
            metrics.standardLactationMilk || "",
            metrics.peakYield || "",
            metrics.lactationLength || "",
          ];
        }));
      autoTable(doc, {
        startY: 42,
        head: [["Animal No.", "Parity", "Mode", "Calving Date", "Dry Date", "TLMY", "SLMY", "PY", "LL"]],
        body: rows.length ? rows : [["No records", "", "", "", "", "", "", "", ""]],
        theme: "grid",
        headStyles: { fillColor: [16, 120, 83] },
        styles: { fontSize: 7 },
      });
      doc.save(`DairyScope-production-report-${new Date().toISOString().slice(0, 10)}.pdf`);
      return;
    }

    if (reportType === "annual") {
      header("ANNUAL PERFORMANCE REPORT", `Financial Year: ${selectedFinancialYear || "Current"}`);
      const rows = annualFinancialSummary.map((r) => [
        r.parameter || r.label || "",
        ...(r.values || r.months || []),
      ]);
      const fallbackRows = monthlyFinancialSummary.map((r) => [
        r.month,
        r.calvings,
        r.stillbirths,
        r.abortions,
        r.mortality,
        r.lactationsCompleted,
      ]);
      autoTable(doc, {
        startY: 42,
        head: fallbackRows.length ? [["Month", "Calvings", "Stillbirths", "Abortions", "Mortality", "Lactations Completed"]] : [["Parameter", "Value"]],
        body: fallbackRows.length ? fallbackRows : (rows.length ? rows : [["No records", ""]]),
        theme: "grid",
        headStyles: { fillColor: [16, 120, 83] },
        styles: { fontSize: 8 },
      });
      doc.save(`DairyScope-annual-performance-${new Date().toISOString().slice(0, 10)}.pdf`);
    }
  }


  function exportDatabaseBackup() {
    const payload = {
      app: "DairyScope",
      exportedAt: new Date().toISOString(),
      animals,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `DairyScope-backup-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function importDatabaseBackup() {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json,application/json";
    input.onchange = async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        const nextAnimals = Array.isArray(parsed) ? parsed : parsed.animals;
        if (!Array.isArray(nextAnimals)) {
          window.alert("Invalid backup file. No animal records were found.");
          return;
        }
        const ok = window.confirm(`Restore ${nextAnimals.length} animal record(s)? This will replace the current local database.`);
        if (!ok) return;
        setAnimals(nextAnimals.map((animal) => withDefaults(animal)));
        setSelectedId(null);
        setSidebarView("home");
        window.alert("Backup restored successfully.");
      } catch (error) {
        window.alert("Backup restore failed. Please select a valid DairyScope JSON backup.");
      }
    };
    input.click();
  }


  function exportHomeDecisionPdf() {
    const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
    doc.setFontSize(16);
    doc.text("DairyScope Decision Dashboard", 40, 42);
    doc.setFontSize(10);
    doc.text(`Generated: ${formatDateDisplay(new Date())}`, 40, 60);

    autoTable(doc, {
      startY: 82,
      head: [["KPI", "Value", "Note"]],
      body: homeKpis.map((row) => [row.label, String(row.value), row.note]),
      theme: "grid",
      headStyles: { fillColor: [16, 120, 83] },
    });

    let y = doc.lastAutoTable.finalY + 24;
    doc.setFontSize(13);
    doc.text("Alerts", 40, y);

    const alertRows = homeDecisionSupport.alerts.flatMap((group) =>
      (group.items.length ? group.items : [{ tag: "Nil", detail: "No current alert" }]).map((item) => [group.title, item.tag, item.detail])
    );

    autoTable(doc, {
      startY: y + 12,
      head: [["Alert", "Animal Tag No.", "Detail"]],
      body: alertRows,
      theme: "grid",
      headStyles: { fillColor: [16, 120, 83] },
    });

    doc.save(`DairyScope-decision-dashboard-${new Date().toISOString().slice(0, 10)}.pdf`);
  }


  function addAnimal() {
    if (!newAnimal.tagNo.trim()) return;
    const duplicateTag = animals.some((animal) => String(animal.tagNo || "").trim().toLowerCase() === String(newAnimal.tagNo || "").trim().toLowerCase() && animal.id !== editAnimalId);
    if (duplicateTag) {
      console.warn(`Data quality flag: duplicate animal tag number ${newAnimal.tagNo}. Entry was allowed.`);
    }
    const prepared = normalizeAnimalFormData(newAnimal);
    if (editAnimalId) {
      setAnimals((prev) => {
        let updatedAnimal = null;
        const mapped = prev.map((a) => {
          if (a.id !== editAnimalId) return a;
          updatedAnimal = withDefaults({ ...a, ...prepared, id: a.id });
          return updatedAnimal;
        });
        return updatedAnimal?.category === "Female" ? syncDamCalvesInHerd(mapped, updatedAnimal) : mapped;
      });
      setSelectedId(editAnimalId);
      setEditAnimalId(null);
    } else {
      const item = withDefaults({ id: Date.now(), ...prepared });
      setAnimals((prev) => [item, ...prev]);
      setSelectedId(item.id);
    }
    setNewAnimal({ ...emptyAnimal });
    setShowAdd(false);
    setSidebarView("current");
  }

  function startEditAnimal() {
    if (!activeAnimal) return;
    setEditAnimalId(activeAnimal.id);
    setNewAnimal({
      tagNo: activeAnimal.tagNo || "",
      breed: activeAnimal.breed || "Nili-Ravi buffalo",
      dob: activeAnimal.dob || "",
      category: activeAnimal.category || "Female",
      identificationMark: activeAnimal.identificationMark || "",
      status: activeAnimal.status || "Active (present in herd)",
      entrySource: activeAnimal.entrySource || "Already part of herd",
      entryDate: activeAnimal.entryDate || "",
      exitDate: activeAnimal.exitDate || "",
      exitReason: activeAnimal.exitReason || "",
      isBreedingBull: activeAnimal.isBreedingBull || "No",
      breedingSet: activeAnimal.breedingSet || "",
      currentSet: activeAnimal.currentSet || "Yes",
      isPtBull: activeAnimal.isPtBull || "No",
    });
    setShowAdd(true);
    setSidebarView("add");
  }

  function deleteSelectedAnimal() {
    if (!uiSelectedAnimal) return;
    const ok = window.confirm(`Delete animal tag no. ${uiSelectedAnimal.tagNo}? This cannot be undone.`);
    if (!ok) return;
    setAnimals((prev) => prev.filter((a) => a.id !== uiSelectedAnimal.id));
    setSelectedId(null);
    setDetailTab("pedigree");
    setEditAnimalId(null);
    setShowAdd(false);
    setSidebarView(sidebarView === "archives" ? "archives" : "current");
  }

  function cancelForm() {
    setShowAdd(false);
    setEditAnimalId(null);
    setNewAnimal({ ...emptyAnimal });
    setSidebarView("");
  }

  function patchSelected(fn) {
    setAnimals((prev) => {
      let updatedSelected = null;
      const mapped = prev.map((a) => {
        if (a.id !== selectedId) return a;
        updatedSelected = fn(withDefaults(a));
        return updatedSelected;
      });
      return updatedSelected?.category === "Female" ? syncDamCalvesInHerd(mapped, updatedSelected) : mapped;
    });
  }

  function updateFemalePedigree(key, value) {
    patchSelected((a) => ({ ...a, femaleDetails: { ...a.femaleDetails, pedigree: { ...a.femaleDetails.pedigree, [key]: value } } }));
  }

  function updateMalePedigree(key, value) {
    patchSelected((a) => ({ ...a, maleDetails: { ...a.maleDetails, pedigree: { ...a.maleDetails.pedigree, [key]: value } } }));
  }

  function getSelectedPhotoList() {
    if (!activeAnimal) return [];
    return activeAnimal.category === "Female" ? (activeAnimal.femaleDetails?.photographs || []) : (activeAnimal.maleDetails?.photographs || []);
  }

  function updateAnimalPhotographs(nextPhotos) {
    patchSelected((a) => {
      const cleanPhotos = (nextPhotos || []).slice(0, 10);
      if (a.category === "Female") {
        return { ...a, femaleDetails: { ...a.femaleDetails, photographs: cleanPhotos } };
      }
      if (a.category === "Male") {
        return { ...a, maleDetails: { ...a.maleDetails, photographs: cleanPhotos } };
      }
      return a;
    });
  }

  function handleAnimalPhotoUpload(event) {
    const files = Array.from(event.target.files || []);
    event.target.value = "";
    if (!files.length || !activeAnimal) return;
    const existing = getSelectedPhotoList();
    const availableSlots = Math.max(0, 10 - existing.length);
    if (!availableSlots) {
      window.alert("Maximum 10 photographs can be saved for one animal.");
      return;
    }
    const allowedTypes = new Set(["image/jpeg", "image/png", "image/heic", "image/heif"]);
    const allowedExt = /\.(jpe?g|png|heic|heif)$/i;
    const selectedFiles = files.filter((file) => allowedTypes.has(file.type) || allowedExt.test(file.name || "")).slice(0, availableSlots);
    if (!selectedFiles.length) {
      window.alert("Only JPEG, PNG and HEIC photographs are allowed.");
      return;
    }
    Promise.all(selectedFiles.map((file) => new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve({
        id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        name: file.name || "animal-photo",
        type: file.type || (file.name?.toLowerCase().endsWith(".png") ? "image/png" : file.name?.toLowerCase().match(/\.hei[cf]$/) ? "image/heic" : "image/jpeg"),
        size: file.size || 0,
        dataUrl: reader.result,
        addedAt: new Date().toISOString(),
      });
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(file);
    }))).then((newPhotos) => {
      const validPhotos = newPhotos.filter(Boolean);
      updateAnimalPhotographs([...existing, ...validPhotos].slice(0, 10));
    });
    if (files.length > availableSlots) {
      window.alert(`Only ${availableSlots} more photograph(s) can be added. Maximum limit is 10.`);
    }
  }

  function removeAnimalPhoto(photoId) {
    const existing = getSelectedPhotoList();
    updateAnimalPhotographs(existing.filter((photo) => photo.id !== photoId));
  }

  function syncParityAfterCalvingEvents(a, parities, changedParityNo) {
    let reproductionParities = [...(a.femaleDetails.reproductionParities || [])];
    let selectedReproParity = a.femaleDetails.selectedReproParity || "0";
    const calvingParityNo = Number(changedParityNo || 0);
    const changedParity = parities.find((p) => Number(p.parityNo) === calvingParityNo);

    const hasCompleteProductiveOutcome = getCalvingOutcomeEvents(changedParity).some((event) => isCompleteProductiveCalvingEvent(event));

    if (hasCompleteProductiveOutcome) {
      // Productive calving closes the previous reproductive parity and opens the next one.
      const nextReproParityNo = String(calvingParityNo);
      if (!reproductionParities.some((p) => String(p.parityNo) === nextReproParityNo)) {
        reproductionParities.push(makeReproParity(nextReproParityNo));
      }
      selectedReproParity = nextReproParityNo;

      // Next calving parity is intentionally not opened here.
      // It will appear only after the next reproductive parity has insemination/conception details
      // and the user clicks Next Tab from the Reproduction tab.
    }

    const nextAnimal = withDefaults({
      ...a,
      femaleDetails: {
        ...a.femaleDetails,
        calvingParities: parities.sort((x, y) => Number(x.parityNo) - Number(y.parityNo)),
        reproductionParities: reproductionParities.sort((x, y) => Number(x.parityNo) - Number(y.parityNo)),
        selectedReproParity,
      },
    });

    return syncProductionLactationsWithCalvingOutcomes(nextAnimal);
  }

  function finalizeCalvingParityBeforeNextTab() {
    if (!activeAnimal || activeAnimal.category !== "Female") return;
    setAnimals((prev) => prev.map((item) => {
      if (item.id !== activeAnimal.id) return item;
      const a = withDefaults(item);
      let changedParityNo = null;
      const parities = (a.femaleDetails.calvingParities || []).map((cp) => {
        const events = getCalvingOutcomeEvents(cp);
        const updatedEvents = events.map((event) => {
          if (!event.finalized && isReadyToFinalizeProductiveCalvingEvent(event)) {
            changedParityNo = cp.parityNo;
            return { ...event, finalized: true };
          }
          return event;
        });
        return { ...cp, outcomeEvents: updatedEvents };
      });
      if (!changedParityNo) return a;
      return syncParityAfterCalvingEvents(a, parities, changedParityNo);
    }));
  }

  function prepareCalvingParityFromReproBeforeNextTab() {
    if (!activeAnimal || activeAnimal.category !== "Female") return;
    const selectedParityNo = String(activeAnimal.femaleDetails?.selectedReproParity || "");
    if (selectedParityNo === "") return;
    const reproParity = getReproParityByNo(activeAnimal, selectedParityNo);
    const hasConceptionOrConceivedAI = Boolean(getConceptionDateForReproParity(reproParity));
    if (!hasConceptionOrConceivedAI) return;

    const targetCalvingParityNo = Number(selectedParityNo) + 1;
    if (targetCalvingParityNo <= 0) return;

    setAnimals((prev) => prev.map((item) => {
      if (item.id !== activeAnimal.id) return item;
      const a = withDefaults(item);
      const exists = (a.femaleDetails.calvingParities || []).some((cp) => Number(cp.parityNo) === targetCalvingParityNo);
      if (exists) return a;
      return withDefaults({
        ...a,
        femaleDetails: {
          ...a.femaleDetails,
          calvingParities: [...a.femaleDetails.calvingParities, makeCalvingParity(targetCalvingParityNo)].sort((x, y) => Number(x.parityNo) - Number(y.parityNo)),
        },
      });
    }));
  }

  function goToNextFemaleTab() {
    if (detailTab === "calving") finalizeCalvingParityBeforeNextTab();
    if (detailTab === "reproduction") prepareCalvingParityFromReproBeforeNextTab();
    setDetailTab(nextDetailTab(detailTab, FEMALE_TABS));
  }

  function updateCalvingOutcomeEvent(parityIdx, outcomeIdx, key, value) {
    const existingParity = activeAnimal?.femaleDetails?.calvingParities?.[parityIdx];
    const existingEvent = getCalvingOutcomeEvents(existingParity)[outcomeIdx] || makeCalvingOutcomeEvent(outcomeIdx + 1);

    if (key === "calvingDate" && value && activeAnimal?.category === "Female") {
      const selectedParityNo = existingParity?.parityNo;
      const conceptionDate = getConceptionDateForOutcome(activeAnimal, selectedParityNo, outcomeIdx);
      if (conceptionDate) {
        const gp = daysBetween(conceptionDate, value);
        if (gp < 260 || gp > 340) {
          console.warn(`Data quality flag: abnormal gestation/pregnancy length (${gp} days). Entry was allowed.`);
        }
        if (parseDisplayDate(value) && parseDisplayDate(conceptionDate) && parseDisplayDate(value).getTime() < parseDisplayDate(conceptionDate).getTime()) {
          console.warn("Data quality flag: outcome date before conception/AI date. Entry was allowed.");
        }
      }
    }

    if (key === "calfTag" && value) {
      const duplicateCalf = animals.some((animal) => String(animal.tagNo || "").trim().toLowerCase() === String(value || "").trim().toLowerCase());
      if (duplicateCalf) console.warn(`Data quality flag: calf tag ${value} already exists. Entry was allowed.`);
    }

    patchSelected((a) => {
      const parities = [...a.femaleDetails.calvingParities];
      const parity = { ...parities[parityIdx] };
      const outcomeEvents = getCalvingOutcomeEvents(parity);
      const updatedEvent = { ...(outcomeEvents[outcomeIdx] || makeCalvingOutcomeEvent(outcomeIdx + 1)), [key]: value };

      if (key === "calvingDate" && value) {
        updatedEvent.calfSire = getCalfSireForCalving(a, parity.parityNo, outcomeIdx);
        a = { ...a, preCalvingLifecycle: getFemaleLifecycle(a) };
      }

      // Abortion and stillbirth do not create a live calf identity record.
      if (key === "calvingOutcome" && !isCalfIdentityOutcome(updatedEvent)) {
        updatedEvent.calfSex = "";
        updatedEvent.calfTag = "";
        updatedEvent.calfSire = "";
      }

      if (key === "calvingOutcome" && isCalfIdentityOutcome(updatedEvent) && updatedEvent.calvingDate) {
        updatedEvent.calfSire = updatedEvent.calfSire || getCalfSireForCalving(a, parity.parityNo, outcomeIdx);
      }

      outcomeEvents[outcomeIdx] = updatedEvent;
      parity.outcomeEvents = outcomeEvents.map((event, idx) => ({ ...event, outcomeNo: String(idx + 1) }));

      // Preserve first event on legacy top-level fields for backward compatibility.
      const first = parity.outcomeEvents[0] || makeCalvingOutcomeEvent(1);
      parity.calvingDate = first.calvingDate || "";
      parity.calfSex = first.calfSex || "";
      parity.calfTag = first.calfTag || "";
      parity.calfSire = first.calfSire || "";
      parity.calvingOutcome = first.calvingOutcome || "Normal calving";
      parity.remarks = first.remarks || "";

      parities[parityIdx] = parity;
      return withDefaults({ ...a, femaleDetails: { ...a.femaleDetails, calvingParities: parities } });
    });
  }

  // Backward-compatible wrapper used by any old references.
  function updateCalvingParity(rowIdx, key, value) {
    return updateCalvingOutcomeEvent(rowIdx, 0, key, value);
  }

  function addCalvingOutcomeEvent(parityIdx) {
    patchSelected((a) => {
      const parities = [...a.femaleDetails.calvingParities];
      const parity = { ...parities[parityIdx] };
      const outcomeEvents = getCalvingOutcomeEvents(parity);
      outcomeEvents.push(makeCalvingOutcomeEvent(outcomeEvents.length + 1));
      parity.outcomeEvents = outcomeEvents;
      parities[parityIdx] = parity;
      return withDefaults({ ...a, femaleDetails: { ...a.femaleDetails, calvingParities: parities } });
    });
  }

  function removeCalvingOutcomeEvent(parityIdx) {
    patchSelected((a) => {
      const parities = [...a.femaleDetails.calvingParities];
      const parity = { ...parities[parityIdx] };
      const outcomeEvents = getCalvingOutcomeEvents(parity);
      if (outcomeEvents.length <= 1) return a;
      parity.outcomeEvents = outcomeEvents.slice(0, -1).map((event, idx) => ({ ...event, outcomeNo: String(idx + 1) }));
      const first = parity.outcomeEvents[0] || makeCalvingOutcomeEvent(1);
      parity.calvingDate = first.calvingDate || "";
      parity.calfSex = first.calfSex || "";
      parity.calfTag = first.calfTag || "";
      parity.calfSire = first.calfSire || "";
      parity.calvingOutcome = first.calvingOutcome || "Normal calving";
      parity.remarks = first.remarks || "";
      parities[parityIdx] = parity;
      return withDefaults({ ...a, femaleDetails: { ...a.femaleDetails, calvingParities: parities } });
    });
  }


  function incrementCalvingParity() {
    patchSelected((a) => syncProductionLactationsWithCalvingOutcomes(withDefaults({ ...a, femaleDetails: { ...a.femaleDetails, calvingParities: [...a.femaleDetails.calvingParities, makeCalvingParity(a.femaleDetails.calvingParities.length + 1)] } })));
  }

  function decrementCalvingParity() {
    patchSelected((a) => {
      if (a.femaleDetails.calvingParities.length <= 1) return a;
      return syncProductionLactationsWithCalvingOutcomes(withDefaults({ ...a, femaleDetails: { ...a.femaleDetails, calvingParities: a.femaleDetails.calvingParities.slice(0, -1) } }));
    });
  }

  function selectReproParity(parityNo) {
    patchSelected((a) => ({ ...a, femaleDetails: { ...a.femaleDetails, selectedReproParity: String(parityNo) } }));
  }

  function incrementReproParity() {
    patchSelected((a) => {
      const current = Number(a.femaleDetails.selectedReproParity || 0);
      const nextNo = String(current + 1);
      const exists = a.femaleDetails.reproductionParities.some((p) => String(p.parityNo) === nextNo);
      return withDefaults({ ...a, femaleDetails: { ...a.femaleDetails, selectedReproParity: nextNo, reproductionParities: exists ? a.femaleDetails.reproductionParities : [...a.femaleDetails.reproductionParities, makeReproParity(nextNo)] } });
    });
  }

  function decrementReproParity() {
    patchSelected((a) => ({ ...a, femaleDetails: { ...a.femaleDetails, selectedReproParity: String(Math.max(0, Number(a.femaleDetails.selectedReproParity || 0) - 1)) } }));
  }

  function isSelectedReproLockedRuntime(animal) {
    if (!animal || animal.category !== "Female") return false;
    const selectedParityNo = String(animal.femaleDetails?.selectedReproParity || "0");
    const editKey = `${animal.id}:repro:${selectedParityNo}`;
    return isReproParityClosedByProductiveCalving(animal, selectedParityNo) && !closedParityEdit[editKey];
  }

  function updateSelectedRepro(key, value) {
    if (isSelectedReproLockedRuntime(activeAnimal)) return;
    patchSelected((a) => {
      const parities = [...a.femaleDetails.reproductionParities];
      const idx = parities.findIndex((p) => String(p.parityNo) === String(a.femaleDetails.selectedReproParity));
      if (idx < 0) return a;
      const updated = { ...parities[idx], [key]: value };
      parities[idx] = key === "conceptionDate" ? parities[idx] : syncReproConceptionFields(updated);
      return { ...a, femaleDetails: { ...a.femaleDetails, reproductionParities: parities } };
    });
  }

  function addAIRecord() {
    if (isSelectedReproLockedRuntime(activeAnimal)) return;
    patchSelected((a) => {
      const parities = [...a.femaleDetails.reproductionParities];
      const idx = parities.findIndex((p) => String(p.parityNo) === String(a.femaleDetails.selectedReproParity));
      if (idx < 0) return a;
      parities[idx] = syncReproConceptionFields({ ...parities[idx], aiRecords: [...parities[idx].aiRecords, { aiDate: "", aiBullNo: "", aiSetNo: "", result: "Pending" }] });
      return { ...a, femaleDetails: { ...a.femaleDetails, reproductionParities: parities } };
    });
  }

  function removeAIRecord() {
    if (isSelectedReproLockedRuntime(activeAnimal)) return;
    patchSelected((a) => {
      const parities = [...a.femaleDetails.reproductionParities];
      const idx = parities.findIndex((p) => String(p.parityNo) === String(a.femaleDetails.selectedReproParity));
      if (idx < 0 || parities[idx].aiRecords.length === 0) return a;
      parities[idx] = syncReproConceptionFields({ ...parities[idx], aiRecords: parities[idx].aiRecords.slice(0, -1) });
      return { ...a, femaleDetails: { ...a.femaleDetails, reproductionParities: parities } };
    });
  }

  function updateAIRecord(rowIdx, key, value) {
    if (isSelectedReproLockedRuntime(activeAnimal)) return;
    if (key === "aiDate" && value && activeAnimal?.category === "Female") {
      const duplicateAI = (activeAnimal.femaleDetails?.reproductionParities || []).some((parity) =>
        (parity.aiRecords || []).some((record, idx) =>
          normalizeDisplayDate(record.aiDate || "") === normalizeDisplayDate(value) &&
          !(String(parity.parityNo) === String(activeAnimal.femaleDetails?.selectedReproParity) && idx === rowIdx)
        )
      );
      if (duplicateAI) {
        console.warn("Data quality flag: duplicate AI date detected. Entry was allowed.");
      }
    }
    patchSelected((a) => {
      const parities = [...a.femaleDetails.reproductionParities];
      const idx = parities.findIndex((p) => String(p.parityNo) === String(a.femaleDetails.selectedReproParity));
      if (idx < 0) return a;
      const aiRecords = [...parities[idx].aiRecords];
      const rec = { ...(aiRecords[rowIdx] || { aiDate: "", aiBullNo: "", aiSetNo: "", result: "Pending" }), [key]: key === "aiBullNo" || key === "aiSetNo" ? String(value || "").trim() : value };
      aiRecords[rowIdx] = rec;
      const updated = syncReproConceptionFields({ ...parities[idx], aiRecords });
      parities[idx] = updated;
      return { ...a, femaleDetails: { ...a.femaleDetails, reproductionParities: parities } };
    });
  }

  function applyBullToSelectedParity(bull) {
    if (isSelectedReproLockedRuntime(activeAnimal)) return;
    patchSelected((a) => {
      const parities = [...a.femaleDetails.reproductionParities];
      const idx = parities.findIndex((p) => String(p.parityNo) === String(a.femaleDetails.selectedReproParity));
      if (idx < 0) return a;

      const aiRecords = [...(parities[idx].aiRecords || [])];
      const baseRecord = aiRecords.length > 0
        ? aiRecords[aiRecords.length - 1]
        : { aiDate: "", aiBullNo: "", aiSetNo: "", result: "Pending" };

      const updatedRecord = {
        ...baseRecord,
        aiBullNo: bull.tagNo || "",
        aiSetNo: bull.breedingSet || "",
      };

      if (aiRecords.length === 0) aiRecords.push(updatedRecord);
      else aiRecords[aiRecords.length - 1] = updatedRecord;

      parities[idx] = { ...parities[idx], aiRecords };
      return { ...a, femaleDetails: { ...a.femaleDetails, reproductionParities: parities } };
    });
  }

  function selectProductionParity(value) {
    patchSelected((a) => ({ ...a, femaleDetails: { ...a.femaleDetails, selectedProductionParity: String(value) } }));
  }

  function addFridayRecord() {
    patchSelected((a) => {
      const lactations = [...a.femaleDetails.productionLactations];
      const idx = lactations.findIndex((l) => String(l.parityNo) === String(a.femaleDetails.selectedProductionParity));
      if (idx < 0) return a;
      const nextDate = getNextFridayRecordDate(lactations[idx]);
      lactations[idx] = { ...lactations[idx], fridayRecords: [...(lactations[idx].fridayRecords || []), makeFridayRecord(nextDate)] };
      return { ...a, femaleDetails: { ...a.femaleDetails, productionLactations: lactations } };
    });
  }

  function removeFridayRecord() {
    patchSelected((a) => {
      const lactations = [...a.femaleDetails.productionLactations];
      const idx = lactations.findIndex((l) => String(l.parityNo) === String(a.femaleDetails.selectedProductionParity));
      if (idx < 0) return a;
      lactations[idx] = { ...lactations[idx], fridayRecords: (lactations[idx].fridayRecords || []).slice(0, -1) };
      return { ...a, femaleDetails: { ...a.femaleDetails, productionLactations: lactations } };
    });
  }

  function updateSelectedLactation(key, value) {
    patchSelected((a) => {
      const lactations = [...a.femaleDetails.productionLactations];
      const idx = lactations.findIndex((l) => String(l.parityNo) === String(a.femaleDetails.selectedProductionParity));
      if (idx < 0) return a;
      const updated = { ...lactations[idx], [key]: value };
      if (key === "entryMode" && value === "Friday Records" && !(updated.fridayRecords || []).length) updated.fridayRecords = [];
      lactations[idx] = updated;
      return { ...a, femaleDetails: { ...a.femaleDetails, productionLactations: lactations } };
    });
  }

  function updateManualSummary(key, value) {
    patchSelected((a) => {
      const lactations = [...a.femaleDetails.productionLactations];
      const idx = lactations.findIndex((l) => String(l.parityNo) === String(a.femaleDetails.selectedProductionParity));
      if (idx < 0) return a;
      lactations[idx] = { ...lactations[idx], manualSummary: { ...lactations[idx].manualSummary, [key]: value } };
      return { ...a, femaleDetails: { ...a.femaleDetails, productionLactations: lactations } };
    });
  }

  function updateFridayRecord(rowIdx, key, value) {
    patchSelected((a) => {
      const lactations = [...a.femaleDetails.productionLactations];
      const idx = lactations.findIndex((l) => String(l.parityNo) === String(a.femaleDetails.selectedProductionParity));
      if (idx < 0) return a;
      const records = [...(lactations[idx].fridayRecords || [])];
      const existing = records[rowIdx] || makeFridayRecord("");
      records[rowIdx] = recalcFridayRecord({ ...existing, [key]: value });
      lactations[idx] = { ...lactations[idx], fridayRecords: records };
      return { ...a, femaleDetails: { ...a.femaleDetails, productionLactations: lactations } };
    });
  }

  function addHealthRecord(target, section, blankRecord) {
    patchSelected((a) => ({
      ...a,
      [target]: {
        ...a[target],
        health: {
          ...a[target].health,
          [section]: [...(a[target].health[section] || []), blankRecord],
        },
      },
    }));
  }

  function removeHealthRecord(target, section) {
    patchSelected((a) => {
      const current = a[target].health[section] || [];
      if (current.length <= 1) return a;
      return {
        ...a,
        [target]: {
          ...a[target],
          health: {
            ...a[target].health,
            [section]: current.slice(0, -1),
          },
        },
      };
    });
  }

  function updateHealthRecord(target, section, rowIdx, key, value) {
    patchSelected((a) => {
      const current = [...(a[target].health[section] || [])];
      current[rowIdx] = { ...(current[rowIdx] || {}), [key]: value };
      return {
        ...a,
        [target]: {
          ...a[target],
          health: {
            ...a[target].health,
            [section]: current,
          },
        },
      };
    });
  }

  function addDiseaseTest() {
    patchSelected((a) => ({ ...a, maleDetails: { ...a.maleDetails, diseaseTests: [...(a.maleDetails.diseaseTests || []), makeDiseaseTestRecord()] } }));
  }

  function removeDiseaseTest() {
    patchSelected((a) => {
      if ((a.maleDetails.diseaseTests || []).length <= 1) return a;
      return { ...a, maleDetails: { ...a.maleDetails, diseaseTests: a.maleDetails.diseaseTests.slice(0, -1) } };
    });
  }

  function updateDiseaseTest(rowIdx, key, value) {
    patchSelected((a) => {
      const current = [...(a.maleDetails.diseaseTests || [])];
      current[rowIdx] = { ...(current[rowIdx] || {}), [key]: value };
      return { ...a, maleDetails: { ...a.maleDetails, diseaseTests: current } };
    });
  }

  function updateFemaleHistoryMeta(key, value) {
    patchSelected((a) => ({ ...a, femaleDetails: { ...a.femaleDetails, historyMeta: { ...a.femaleDetails.historyMeta, [key]: value } } }));
  }

  function updateMaleHistoryMeta(key, value) {
    patchSelected((a) => ({ ...a, maleDetails: { ...a.maleDetails, historyMeta: { ...a.maleDetails.historyMeta, [key]: value } } }));
  }

  const femaleTarget = "femaleDetails";
  const maleTarget = "maleDetails";
  const selectedHealthTarget = activeAnimal?.category === "Female" ? femaleTarget : maleTarget;
  const selectedHealth = activeAnimal?.category === "Female" ? selectedAnimal?.femaleDetails?.health : selectedAnimal?.maleDetails?.health;

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top_left,_rgba(16,185,129,0.12),_transparent_28%),radial-gradient(circle_at_top_right,_rgba(6,182,212,0.12),_transparent_24%),linear-gradient(180deg,_#f7faf9_0%,_#edf7f5_45%,_#eaf4f3_100%)] p-4 text-slate-800 md:p-6">
      <div className="mx-auto max-w-[1780px]">
        <div className="grid gap-6 xl:grid-cols-[270px_1fr]">
          <aside className="hidden xl:flex xl:min-h-[calc(100vh-3rem)] xl:w-[285px] xl:flex-col xl:justify-between rounded-[34px] border border-white/10 bg-slate-950 p-5 text-white shadow-[0_30px_80px_rgba(15,23,42,0.28)]">
            <div>
              <div className="rounded-[28px] border border-white/10 bg-white/[0.04] p-4">
                <div className="flex items-center gap-3">
                  <div className="flex h-16 w-16 items-center justify-center overflow-hidden rounded-full border border-emerald-300/30 bg-white p-2 shadow-[0_12px_28px_rgba(16,185,129,0.18)]">
                    <img src="/logo.png" alt="DairyScope logo" className="h-full w-full object-contain" />
                  </div>
                  <div>
                    <div className="text-xl font-bold tracking-tight">DairyScope</div>
                    <div className="mt-0.5 text-xs leading-5 text-slate-400">Smart Eyes for Your Smart Herd</div>
                  </div>
                </div>
              </div>

              <div className="mt-7 space-y-1.5">
                <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "home" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("home")}>
                  <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "home" ? "bg-emerald-400" : "bg-transparent"}`} />
                  <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">⌂</span>
                  <span>Home</span>
                </button>

                <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "dashboard" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("dashboard")}>
                  <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "dashboard" ? "bg-emerald-400" : "bg-transparent"}`} />
                  <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">▦</span>
                  <span>Herd Structure</span>
                </button>

                <div className="space-y-1.5">
                  <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "current" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("current")}>
                    <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "current" ? "bg-emerald-400" : "bg-transparent"}`} />
                    <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">🐃</span>
                    <span>Current Herd</span>
                  </button>
                  {sidebarView === "current" && (
                    <div className="ml-10 space-y-2 border-l border-white/10 pl-3">
                      <button className={`w-full rounded-xl px-3 py-2 text-left text-xs font-semibold transition ${currentHerdSex === "Female" ? "bg-emerald-500 text-white" : "text-slate-300 hover:bg-white/10"}`} onClick={() => setCurrentHerdSex("Female")}>Female</button>
                      <button className={`w-full rounded-xl px-3 py-2 text-left text-xs font-semibold transition ${currentHerdSex === "Male" ? "bg-emerald-500 text-white" : "text-slate-300 hover:bg-white/10"}`} onClick={() => setCurrentHerdSex("Male")}>Male</button>
                    </div>
                  )}
                </div>

                <div className="space-y-1.5">
                  <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "archives" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("archives")}>
                    <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "archives" ? "bg-emerald-400" : "bg-transparent"}`} />
                    <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">◴</span>
                    <span>Archives</span>
                  </button>
                  {sidebarView === "archives" && (
                    <div className="ml-10 space-y-2 border-l border-white/10 pl-3">
                      <button className={`w-full rounded-xl px-3 py-2 text-left text-xs font-semibold transition ${archiveHerdSex === "Female" ? "bg-emerald-500 text-white" : "text-slate-300 hover:bg-white/10"}`} onClick={() => setArchiveHerdSex("Female")}>Female</button>
                      <button className={`w-full rounded-xl px-3 py-2 text-left text-xs font-semibold transition ${archiveHerdSex === "Male" ? "bg-emerald-500 text-white" : "text-slate-300 hover:bg-white/10"}`} onClick={() => setArchiveHerdSex("Male")}>Male</button>
                    </div>
                  )}
                </div>

                <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "analytics" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("analytics")}>
                  <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "analytics" ? "bg-emerald-400" : "bg-transparent"}`} />
                  <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">⌁</span>
                  <span>Herd Analytics</span>
                </button>

                <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "testday" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("testday")}>
                  <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "testday" ? "bg-emerald-400" : "bg-transparent"}`} />
                  <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">☷</span>
                  <span className="flex-1">Test day Records</span>
                  <span className="rounded-full bg-white/15 px-2 py-0.5 text-[10px] font-bold">WA/HA</span>
                </button>

                <button className={`group relative flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-semibold transition ${sidebarView === "reports" ? "bg-white text-slate-950 shadow-lg" : "text-slate-300 hover:bg-white/10 hover:text-white"}`} onClick={() => setSidebarView("reports")}>
                  <span className={`absolute left-0 top-2 h-8 w-1 rounded-r-full ${sidebarView === "reports" ? "bg-emerald-400" : "bg-transparent"}`} />
                  <span className="grid h-8 w-8 place-items-center rounded-xl bg-white/10 text-base">▣</span>
                  <span>Reports</span>
                </button>

                <button className={`mt-4 flex w-full items-center justify-center gap-2 rounded-2xl bg-gradient-to-r from-emerald-500 to-cyan-500 px-4 py-3 text-sm font-bold text-white shadow-lg transition hover:opacity-95 ${sidebarView === "add" ? "ring-2 ring-white/40" : ""}`} onClick={() => { setSidebarView("add"); setEditAnimalId(null); setNewAnimal({ ...emptyAnimal }); }}>
                  <span>＋</span>
                  <span>Add Animal</span>
                </button>
              </div>
            </div>

            <div className="space-y-3 rounded-[26px] border border-white/10 bg-white/[0.04] p-4 text-xs text-slate-300">
              <div className="flex items-center justify-between">
                <span className="text-slate-400">Version</span>
                <span className="rounded-full bg-emerald-400/10 px-2.5 py-1 font-semibold text-emerald-200">v1.0</span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-slate-400">Institute</span>
                <span className="font-semibold text-white">ICAR-CIRB</span>
              </div>
              <div className="rounded-2xl bg-slate-900/80 p-3">
                <div className="flex items-center gap-2">
                  <span className="h-2.5 w-2.5 rounded-full bg-emerald-400 shadow-[0_0_12px_rgba(52,211,153,0.9)]" />
                  <span className="font-semibold text-slate-100">Backup status</span>
                </div>
                <div className="mt-1 text-[11px] leading-5 text-slate-400">Local backup/export tools available on Home.</div>
              </div>
            </div>
          </aside>

          <div className="space-y-6">
            <div className="rounded-[28px] border border-slate-200/70 bg-white/95 p-4 shadow-[0_8px_22px_rgba(15,23,42,0.05)]">
              <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                <div className="flex items-center gap-4">
                  <div className="hidden h-12 w-12 items-center justify-center overflow-hidden rounded-full border border-emerald-100 bg-white p-1 shadow-sm sm:flex">
                    <img src="/logo.png" alt="DairyScope" className="h-full w-full object-contain" />
                  </div>
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-[0.20em] text-emerald-700">Current selected herd</div>
                    <div className="mt-1 flex flex-wrap items-center gap-2">
                      <span className="text-xl font-bold tracking-tight text-slate-950">{currentHerdSex} · {sidebarView === "archives" ? "Archives" : "Active Herd"}</span>
                      <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-700">{activeAnimals.length} active</span>
                    </div>
                  </div>
                </div>

                <div className="flex flex-1 flex-col gap-3 xl:max-w-3xl xl:flex-row xl:items-center xl:justify-end">
                  <div className="relative flex-1 xl:max-w-sm">
                    <span className="pointer-events-none absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">⌕</span>
                    <input
                      className="w-full rounded-2xl border border-slate-200 bg-slate-50/80 px-11 py-3 text-sm font-medium text-slate-800 shadow-inner outline-none transition focus:border-emerald-300 focus:bg-white focus:ring-4 focus:ring-emerald-100"
                      placeholder="Search animal tag, breed, status..."
                      value={search}
                      onChange={(e) => setSearch(e.target.value)}
                    />
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <button className="rounded-2xl border border-emerald-100 bg-white px-4 py-3 text-xs font-bold uppercase tracking-[0.12em] text-emerald-700 shadow-sm transition hover:bg-emerald-50" onClick={exportDatabaseBackup}>
                      Backup
                    </button>
                    <button className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-xs font-bold uppercase tracking-[0.12em] text-slate-700 shadow-sm transition hover:bg-slate-50" onClick={exportHomeDecisionPdf}>
                      Export
                    </button>
                    <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-xs font-semibold text-slate-600">
                      {todayHeaderLabel}
                    </div>
                    <div className="flex items-center gap-2 rounded-2xl bg-slate-950 px-4 py-3 text-xs font-semibold text-white shadow-sm">
                      <span className="grid h-6 w-6 place-items-center rounded-full bg-emerald-500 text-[10px]">DS</span>
                      <span>Farm User</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>



        {sidebarView === "reports" && (
          <Section title="Reports">
            <div className="space-y-6">
              <div className="rounded-[28px] border border-emerald-100 bg-emerald-50/60 p-6">
                <div className="text-2xl font-bold tracking-tight text-slate-950">Professional Reporting Centre</div>
                <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">
                  Generate official PDF reports for individual animals, herd structure, fertility performance, production and annual institutional review.
                </p>
              </div>

              <div className="grid gap-5 md:grid-cols-2 xl:grid-cols-3">
                <button className="rounded-[26px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] transition hover:-translate-y-1 hover:shadow-lg" onClick={() => exportDairyScopeReport("femaleHistory")}>
                  <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-emerald-50 text-xl">♀</div>
                  <div className="mt-4 text-lg font-bold text-slate-950">Female History Sheet</div>
                  <p className="mt-2 text-sm leading-6 text-slate-600">Export the complete selected female animal history sheet.</p>
                </button>

                <button className="rounded-[26px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] transition hover:-translate-y-1 hover:shadow-lg" onClick={() => exportDairyScopeReport("bullHistory")}>
                  <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-cyan-50 text-xl">♂</div>
                  <div className="mt-4 text-lg font-bold text-slate-950">Breeding Bull History Sheet</div>
                  <p className="mt-2 text-sm leading-6 text-slate-600">Export selected breeding bull history, progeny and daughter performance.</p>
                </button>

                <button className="rounded-[26px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] transition hover:-translate-y-1 hover:shadow-lg" onClick={() => exportDairyScopeReport("herdStructure")}>
                  <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-blue-50 text-xl">▦</div>
                  <div className="mt-4 text-lg font-bold text-slate-950">Herd Structure Report</div>
                  <p className="mt-2 text-sm leading-6 text-slate-600">Summarise active herd, females, males, milk, dry, heifers and bulls.</p>
                </button>

                <button className="rounded-[26px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] transition hover:-translate-y-1 hover:shadow-lg" onClick={() => exportDairyScopeReport("aiFertility")}>
                  <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-violet-50 text-xl">AI</div>
                  <div className="mt-4 text-lg font-bold text-slate-950">AI &amp; Fertility Report</div>
                  <p className="mt-2 text-sm leading-6 text-slate-600">Export AI records, bull/set use and conception result summary.</p>
                </button>

                <button className="rounded-[26px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] transition hover:-translate-y-1 hover:shadow-lg" onClick={() => exportDairyScopeReport("production")}>
                  <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-amber-50 text-xl">🥛</div>
                  <div className="mt-4 text-lg font-bold text-slate-950">Production Report</div>
                  <p className="mt-2 text-sm leading-6 text-slate-600">Export lactation-wise TLMY, SLMY, peak yield and lactation length.</p>
                </button>

                <button className="rounded-[26px] border border-slate-200/70 bg-white/95 p-5 text-left shadow-[0_12px_28px_rgba(15,23,42,0.06)] transition hover:-translate-y-1 hover:shadow-lg" onClick={() => exportDairyScopeReport("annual")}>
                  <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-slate-100 text-xl">FY</div>
                  <div className="mt-4 text-lg font-bold text-slate-950">Annual Performance Report</div>
                  <p className="mt-2 text-sm leading-6 text-slate-600">Export annual financial-year performance summary for review meetings.</p>
                </button>
              </div>
            </div>
          </Section>
        )}


        {sidebarView === "analytics" && (
          <Section title="Analytical Summary">
            <div className="grid gap-5 lg:grid-cols-2">
              <div className="rounded-2xl border border-emerald-100 bg-emerald-50/60 p-4">
                <div className="mb-3 text-xl font-semibold text-emerald-900">Monthly Summary · Financial Year</div>
                <div className="mb-4 grid gap-4 md:grid-cols-2">
                  <SelectField label="Financial Year" value={selectedFinancialYear} onChange={(v) => { setSelectedFinancialYear(v); setSummaryDetail(null); }} options={financialYearOptions} />
                  <SelectField label="Month" value={selectedFinancialMonth} onChange={(v) => { setSelectedFinancialMonth(v); setSummaryDetail(null); }} options={["Overall","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]} />
                </div>
                <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3">
                  <table className="min-w-full border-collapse text-sm">
                    <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Month","Calvings","Stillbirths","Abortions","Mortality","Lactations Completed"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                    <tbody>
                      {monthlyFinancialSummary.map((r) => (
                        <tr key={`fy-${r.month}`}>
                          <td className="border border-emerald-100 px-2 py-2">{r.month}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.calvings > 0 ? <button className="font-semibold text-emerald-700 underline underline-offset-2" onClick={() => setSummaryDetail({ type: "calvings", month: r.month })}>{r.calvings}</button> : r.calvings}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.stillbirths > 0 ? <button className="font-semibold text-emerald-700 underline underline-offset-2" onClick={() => setSummaryDetail({ type: "stillbirths", month: r.month })}>{r.stillbirths}</button> : r.stillbirths}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.abortions > 0 ? <button className="font-semibold text-emerald-700 underline underline-offset-2" onClick={() => setSummaryDetail({ type: "abortions", month: r.month })}>{r.abortions}</button> : r.abortions}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.mortality > 0 ? <button className="font-semibold text-emerald-700 underline underline-offset-2" onClick={() => setSummaryDetail({ type: "mortality", month: r.month })}>{r.mortality}</button> : r.mortality}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.lactationsCompleted > 0 ? <button className="font-semibold text-emerald-700 underline underline-offset-2" onClick={() => setSummaryDetail({ type: "lactationsCompleted", month: r.month })}>{r.lactationsCompleted}</button> : r.lactationsCompleted}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="mt-3 text-sm text-slate-600">Financial-year reporting runs from April to March.</div>
                {summaryDetail && (
                  <SummaryDetailTable
                    detailType={summaryDetail.type}
                    rows={buildMonthlyEventDetails(normalizedAnimals, selectedFinancialYearNum, summaryDetail.month, summaryDetail.type)}
                    monthLabel={summaryDetail.month}
                    financialYearLabel={financialYearOptions.find((o) => String(o.value) === String(selectedFinancialYear))?.label || String(selectedFinancialYear)}
                    onClose={() => setSummaryDetail(null)}
                  />
                )}
              </div>

              <div className="rounded-2xl border border-emerald-100 bg-emerald-50/60 p-4">
                <div className="mb-3 text-xl font-semibold text-emerald-900">Conception Dashboard · Calendar Year</div>
                <div className="mb-4 grid gap-4 md:grid-cols-2">
                  <SelectField label="Calendar Year" value={selectedCalendarYear} onChange={setSelectedCalendarYear} options={calendarYearOptions} />
                  <SelectField label="Month" value={selectedCalendarMonth} onChange={setSelectedCalendarMonth} options={["Overall","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]} />
                </div>
                <div className="mb-4 grid gap-4 md:grid-cols-3">
                  <StatCard title="Total AI" value={annualConceptionSummary.totalAI} />
                  <StatCard title="Conceived" value={annualConceptionSummary.conceived} />
                  <StatCard title="CR %" value={annualConceptionSummary.crPct} />
                </div>
                <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3">
                  <table className="min-w-full border-collapse text-sm">
                    <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Month","Total AI","Total Conceived","Conception Rate %"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                    <tbody>
                      {monthlyConceptionRows.map((r) => (
                        <tr key={`cy-${r.month}`}>
                          <td className="border border-emerald-100 px-2 py-2">{r.totalAI > 0 ? <button className="font-semibold text-emerald-700 underline underline-offset-2" onClick={() => setSelectedConceptionDetailMonth(r.month)}>{r.month}</button> : r.month}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.totalAI}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.conceived}</td>
                          <td className="border border-emerald-100 px-2 py-2">{r.crPct}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="mt-3 text-sm text-slate-600">Conception reporting runs from January to December.</div>
                {selectedConceptionDetailMonth && (
                  <div className="mt-4 rounded-2xl border border-emerald-100 bg-white p-4">
                    <div className="mb-3 flex items-center justify-between gap-3">
                      <div>
                        <div className="text-lg font-semibold text-emerald-900">AI details · {selectedConceptionDetailMonth} {selectedCalendarYearNum}</div>
                        <div className="text-sm text-slate-600">All inseminations recorded during the selected month.</div>
                      </div>
                      <button className="rounded-xl border border-emerald-200 px-3 py-2 text-sm font-medium text-emerald-700 hover:bg-emerald-50" onClick={() => setSelectedConceptionDetailMonth(null)}>Close</button>
                    </div>
                    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                      <table className="min-w-full border-collapse text-sm">
                        <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Animal inseminated","Date of last calving","No. of AI so far","Bull used for current AI","Date of current AI","Result of insemination"].map((h)=><th key={h} className="whitespace-nowrap border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                        <tbody>
                          {conceptionDetailRows.map((row, idx) => (
                            <tr key={`conception-detail-${idx}`}>
                              <td className="border border-emerald-100 px-2 py-2">{row.tagNo}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.lastCalvingDate}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.totalAISoFar}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.bullUsed}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.currentAIDate}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.inseminationResult}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {conceptionDetailRows.length === 0 && <div className="mt-3 text-sm text-slate-500">No AI records found for this month.</div>}
                  </div>
                )}
              </div>
            </div>

            <div className="mt-5 grid gap-5 lg:grid-cols-2">
              <div className="rounded-2xl border border-emerald-100 bg-white p-4">
                <div className="mb-3 text-lg font-semibold text-emerald-900">Quarter-wise Conception Rate</div>
                <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                  <table className="min-w-full border-collapse text-sm">
                    <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Quarter","No. of AI","Pregnant animals","CR %"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                    <tbody>
                      {quarterWiseConceptionRows.map((r)=><tr key={r.quarter}>
                        <td className="border border-emerald-100 px-2 py-2">{r.quarter}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.totalAI}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.conceived}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.crPct}</td>
                      </tr>)}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="rounded-2xl border border-emerald-100 bg-white p-4">
                <div className="mb-3 text-lg font-semibold text-emerald-900">Bull-wise Conception Rate</div>
                <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                  <table className="min-w-full border-collapse text-sm">
                    <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Bull No.","Set No.","Total No. of AI","Total Conceived","CR %"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                    <tbody>
                      {bullWiseConceptionRows.map((r)=><tr key={`${r.bullNo}-${r.setNo}`}>
                        <td className="border border-emerald-100 px-2 py-2">{r.bullNo}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.setNo || "—"}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.totalAI}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.conceived}</td>
                        <td className="border border-emerald-100 px-2 py-2">{r.crPct}</td>
                      </tr>)}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            <div className="mt-5 rounded-2xl border border-emerald-100 bg-white p-4">
              <div className="mb-3 text-lg font-semibold text-emerald-900">Female Conception Rate by AI number (nth)</div>
              <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                <table className="min-w-full border-collapse text-sm">
                  <thead className="bg-slate-50/90">
                    <tr className="bg-emerald-50 text-left text-emerald-900">
                      <th className="border border-emerald-100 px-2 py-2">Group</th>
                      <th className="border border-emerald-100 px-2 py-2">1st AI</th>
                      <th className="border border-emerald-100 px-2 py-2">2nd AI</th>
                      <th className="border border-emerald-100 px-2 py-2">3rd AI</th>
                      <th className="border border-emerald-100 px-2 py-2">4th & above</th>
                    </tr>
                  </thead>
                  <tbody>
                    {["Heifer","Milk+ dry"].map((grp)=><tr key={grp}>
                      <td className="border border-emerald-100 px-2 py-2 font-medium">{grp}</td>
                      {[1,2,3,4].map((nth)=><td key={`${grp}-${nth}`} className="border border-emerald-100 px-2 py-2">{femaleConceptionByNthAI[grp][nth].ai}/{femaleConceptionByNthAI[grp][nth].conceived} ({femaleConceptionByNthAI[grp][nth].crPct}%)</td>)}
                    </tr>)}
                  </tbody>
                </table>
              </div>
            </div>



            <div className="mt-5 flex justify-end">
              <button
                className="rounded-xl bg-emerald-600 px-4 py-2 font-medium text-white shadow-sm transition hover:bg-emerald-700"
                onClick={exportAnalyticsSummaryPdf}
              >
                Export Present View PDF
              </button>
            </div>
          </Section>
        )}

        {sidebarView === "testday" && (
          <Section title="Test day Records">
            <div className="rounded-2xl border border-emerald-100 bg-white p-4">
              <div className="mb-3 flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
                <div>
                  <div className="text-lg font-semibold text-emerald-900">WA / HA Friday Test-Day Entry</div>
                  <div className="text-sm text-slate-600">Enter Morning MY and Evening MY directly against each animal in milk. Total MY auto-sums; Save stores the Friday record into WA/HA and the individual animal production tab for the matching parity.</div>
                </div>
                <div className="grid gap-3 md:grid-cols-[180px_auto] md:items-end">
                  <FridayDatePicker label="Friday date" value={selectedWAFriday} onChange={setSelectedWAFriday} />
                  <button className="rounded-xl bg-emerald-600 px-4 py-2 font-medium text-white hover:bg-emerald-700" onClick={saveWAInlineEntries}>Save Friday Records</button>
                </div>
              </div>
              <div className="mb-4 grid gap-4 md:grid-cols-5">
                <StatCard title="Milk animals" value={waFridaySummary.milkAnimals} />
                <StatCard title="Dry animals" value={waFridaySummary.dryAnimals} />
                <StatCard title="Total milk" value={waFridaySummary.totalMilk} />
                <StatCard title="WA" value={waFridaySummary.wa} />
                <StatCard title="HA" value={waFridaySummary.ha} />
              </div>
              <div className="grid gap-5 lg:grid-cols-2">
                <div>
                  <div className="mb-2 text-sm font-semibold text-emerald-900">Direct milk entry for selected Friday</div>
                  <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                    <table className="min-w-full border-collapse text-sm">
                      <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag No.","Parity","Calving Date","Mor MY","Eve MY","Total MY"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                      <tbody>
                        {waEligibleAnimals.length ? waEligibleAnimals.map((animal) => {
                          const lactation = getLactationForDate(animal, selectedWAFriday) || getCurrentOrLatestLactation(animal);
                          const draft = waEntryRows[animal.id] || { morningMilk: "", eveningMilk: "", totalDailyYield: "" };
                          return <tr key={`wa-eligible-${animal.id}`}>
                            <td className="border border-emerald-100 px-2 py-2 font-semibold text-emerald-900">{animal.tagNo}</td>
                            <td className="border border-emerald-100 px-2 py-2">{lactation?.parityNo || "—"}</td>
                            <td className="border border-emerald-100 px-2 py-2">{lactation?.calvingDate || "—"}</td>
                            <td className="border border-emerald-100 px-2 py-2"><input className="w-24 rounded-lg border border-emerald-100 px-2 py-1" type="number" step="0.1" min="0" inputMode="decimal" value={draft.morningMilk} onChange={(e) => updateWAMilkEntry(animal.id, "morningMilk", e.target.value)} /></td>
                            <td className="border border-emerald-100 px-2 py-2"><input className="w-24 rounded-lg border border-emerald-100 px-2 py-1" type="number" step="0.1" min="0" inputMode="decimal" value={draft.eveningMilk} onChange={(e) => updateWAMilkEntry(animal.id, "eveningMilk", e.target.value)} /></td>
                            <td className="border border-emerald-100 px-2 py-2 font-semibold">{draft.totalDailyYield || "—"}</td>
                          </tr>;
                        }) : <tr><td colSpan={8} className="border border-emerald-100 px-2 py-3 text-slate-500">No milk animals are available for the selected Friday.</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
                <div>
                  <div className="mb-2 text-sm font-semibold text-emerald-900">Saved Friday WA / HA history</div>
                  <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                    <table className="min-w-full border-collapse text-sm">
                      <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Friday","Milk animals","Dry animals","Total milk","WA","HA"].map((h)=><th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                      <tbody>
                        {waHistoryRows.length ? waHistoryRows.slice(0, 20).map((row) => <tr key={`wa-history-${row.date}`}>
                          <td className="border border-emerald-100 px-2 py-2">{row.date}</td>
                          <td className="border border-emerald-100 px-2 py-2">{row.milkAnimals}</td>
                          <td className="border border-emerald-100 px-2 py-2">{row.dryAnimals}</td>
                          <td className="border border-emerald-100 px-2 py-2">{row.totalMilk}</td>
                          <td className="border border-emerald-100 px-2 py-2">{row.wa}</td>
                          <td className="border border-emerald-100 px-2 py-2">{row.ha}</td>
                        </tr>) : <tr><td colSpan={8} className="border border-emerald-100 px-2 py-3 text-slate-500">No saved Friday milk records yet.</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          </Section>
        )}

        {(sidebarView === "add" || showAdd) && (
          <Section title={editAnimalId ? "Edit Animal" : "Add Animal"}>
            <Grid>
              <SelectField label="Breed" value={newAnimal.breed} onChange={(v) => setNewAnimal((s) => ({ ...s, breed: v }))} options={BREEDS} />
              <TextField label="Tag No." value={newAnimal.tagNo} onChange={(v) => setNewAnimal((s) => ({ ...s, tagNo: v }))} />
              <DateField label="Date of birth" value={newAnimal.dob} onChange={(v) => setNewAnimal((s) => ({ ...s, dob: v }))} />
              <SelectField label="Category" value={newAnimal.category} onChange={handleFormCategoryChange} options={SEX_OPTIONS} />
              <TextField label="Identification mark" value={newAnimal.identificationMark} onChange={(v) => setNewAnimal((s) => ({ ...s, identificationMark: v }))} />
              <SelectField label="Status" value={newAnimal.status} onChange={handleFormStatusChange} options={STATUS_OPTIONS} />
              <SelectField label="Entry source" value={newAnimal.entrySource || "Already part of herd"} onChange={(v) => setNewAnimal((s) => normalizeAnimalFormData({ ...s, entrySource: v }))} options={ENTRY_SOURCE_OPTIONS} />
              {(newAnimal.entrySource || "Already part of herd") === "Purchased" && (
                <DateField label="Date of purchase / entry" value={newAnimal.entryDate || ""} onChange={(v) => setNewAnimal((s) => ({ ...s, entryDate: v }))} />
              )}
              {newAnimal.category === "Male" && (
                <>
                  <SelectField label="Selected for breeding" value={newAnimal.isBreedingBull || "No"} onChange={(v) => setNewAnimal((s) => normalizeAnimalFormData({ ...s, isBreedingBull: v }))} options={["No", "Yes"]} />
                  {newAnimal.isBreedingBull === "Yes" && <TextField label="Included as breeding in which set (Roman numerals only)" value={newAnimal.breedingSet || ""} onChange={(v) => setNewAnimal((s) => ({ ...s, breedingSet: normalizeRomanInput(v) }))} />}
                  {newAnimal.isBreedingBull === "Yes" && <SelectField label="Current set" value={newAnimal.currentSet || "Yes"} onChange={(v) => setNewAnimal((s) => normalizeAnimalFormData({ ...s, currentSet: v }))} options={["Yes", "No"]} />}
                  {newAnimal.isBreedingBull === "Yes" && (newAnimal.currentSet || "Yes") === "No" && <SelectField label="PT bull" value={newAnimal.isPtBull || "No"} onChange={(v) => setNewAnimal((s) => normalizeAnimalFormData({ ...s, isPtBull: v }))} options={["No", "Yes"]} />}
                </>
              )}
              {newAnimal.status !== "Active (present in herd)" && (
                <>
                  <DateField label="Date of Death / Culling" value={newAnimal.exitDate || ""} onChange={(v) => setNewAnimal((s) => ({ ...s, exitDate: v }))} />
                  <AreaField label="Reason of Death / Culling" value={newAnimal.exitReason || ""} onChange={(v) => setNewAnimal((s) => ({ ...s, exitReason: v }))} rows={3} />
                </>
              )}
            </Grid>
            <div className="mt-4 flex gap-2">
              <button className="rounded-xl bg-emerald-600 px-4 py-2 font-medium text-white hover:bg-emerald-700" onClick={addAnimal}>{editAnimalId ? "Save Changes" : "Save Animal"}</button>
              <button className="rounded-xl border border-slate-300 bg-white px-4 py-2 hover:bg-slate-50" onClick={cancelForm}>Cancel</button>
            </div>
          </Section>
        )}

        

        
        {sidebarView === "home" && (
          <Section title="Home">
            <div className="space-y-8">
              {/* Corporate landing hero */}
              <div className="relative overflow-hidden rounded-[38px] border border-white/70 bg-gradient-to-br from-slate-950 via-emerald-950 to-slate-900 p-8 text-white shadow-[0_24px_70px_rgba(15,23,42,0.22)] lg:p-10">
                <div className="absolute -right-20 -top-24 h-72 w-72 rounded-full bg-emerald-400/20 blur-3xl" />
                <div className="absolute -bottom-24 left-1/3 h-72 w-72 rounded-full bg-cyan-400/10 blur-3xl" />
                <div className="relative z-10 flex flex-col gap-8 lg:flex-row lg:items-center lg:justify-between">
                  <div className="max-w-4xl">
                    <div className="inline-flex rounded-full border border-white/15 bg-white/10 px-4 py-2 text-xs font-semibold uppercase tracking-[0.24em] text-emerald-100">
                      ICAR-CIRB · Precision Buffalo Intelligence
                    </div>
                    <h1 className="mt-6 text-4xl font-bold leading-tight tracking-tight lg:text-5xl">
                      DairyScope
                    </h1>
                    <p className="mt-4 max-w-3xl text-xl font-semibold leading-8 text-emerald-50">
                      Smart herd intelligence for precision buffalo breeding, fertility monitoring and performance-led decision support.
                    </p>
                    <p className="mt-4 max-w-4xl text-sm leading-7 text-slate-200">
                      DairyScope is an integrated herd intelligence and genetic decision-support platform designed for precision buffalo breeding and performance optimization.
                    </p>
                    <div className="mt-6 flex flex-wrap gap-3 text-xs font-semibold uppercase tracking-[0.16em] text-slate-200">
                      <span className="rounded-full bg-white/10 px-4 py-2">Herd records</span>
                      <span className="rounded-full bg-white/10 px-4 py-2">Breeding decisions</span>
                      <span className="rounded-full bg-white/10 px-4 py-2">Fertility alerts</span>
                      <span className="rounded-full bg-white/10 px-4 py-2">Official reports</span>
                    </div>
                  </div>
                  <div className="flex justify-center lg:justify-end">
                    <div className="flex h-52 w-52 items-center justify-center overflow-hidden rounded-full border border-white/30 bg-white p-4 shadow-[0_22px_50px_rgba(0,0,0,0.28)]">
                      <img src="/logo.png" alt="DairyScope logo" className="h-full w-full object-contain" />
                    </div>
                  </div>
                </div>
              </div>

              {/* KPI strip */}
              <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                {homeKpis.map((kpi) => (
                  <div key={kpi.label} className="rounded-[28px] border border-white/70 bg-white/92 p-5 shadow-[0_14px_32px_rgba(15,23,42,0.07)] transition duration-300 hover:-translate-y-1 hover:shadow-[0_20px_45px_rgba(15,23,42,0.11)]">
                    <div className="flex items-center justify-between gap-4">
                      <div className="rounded-2xl bg-emerald-50 px-3 py-2 text-2xl">{kpi.icon}</div>
                      <div className="text-right text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-400">{kpi.note}</div>
                    </div>
                    <div className="mt-5 text-4xl font-bold tracking-tight text-slate-950">{kpi.value}</div>
                    <div className="mt-1 text-sm font-semibold text-slate-600">{kpi.label}</div>
                  </div>
                ))}
              </div>

              {/* Decision dashboard */}
              <div className="rounded-[34px] border border-white/70 bg-white/92 p-7 shadow-[0_18px_42px_rgba(15,23,42,0.08)]">
                <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-[0.22em] text-emerald-700">Decision Dashboard</div>
                    <h2 className="mt-2 text-3xl font-bold tracking-tight text-slate-950">Management alerts needing attention</h2>
                    <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">
                      Click any alert count to view the animal tag numbers and the exact reason for flagging.
                    </p>
                  </div>
                  <button className="rounded-2xl bg-slate-950 px-5 py-3 text-sm font-semibold text-white shadow-lg hover:bg-slate-800" onClick={exportHomeDecisionPdf}>Export PDF</button>
                </div>

                <div className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
                  {homeDecisionSupport.alerts.map((group) => (
                    <div key={group.id} className="rounded-[24px] border border-slate-200/70 bg-white/90 p-5 shadow-[0_10px_24px_rgba(15,23,42,0.045)] transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                      <div className="flex items-center justify-between gap-3">
                        <div className="font-semibold text-slate-900">{group.title}</div>
                        <button
                          className={`rounded-full px-4 py-1.5 text-sm font-bold transition ${group.items.length ? "bg-rose-50 text-rose-700 hover:bg-rose-100" : "bg-emerald-50 text-emerald-700"}`}
                          onClick={() => group.items.length && setExpandedHomeAlert(expandedHomeAlert === group.id ? "" : group.id)}
                          title={group.items.length ? "Click to view animal tag numbers" : "No animals in this alert"}
                        >
                          {group.items.length}
                        </button>
                      </div>
                      {expandedHomeAlert === group.id ? (
                        <div className="mt-4 space-y-2 text-sm text-slate-600">
                          {group.items.map((item, idx) => (
                            <div key={`${group.id}-${idx}`} className="rounded-2xl border border-slate-100 bg-white px-3 py-2">
                              <span className="font-semibold text-slate-950">{item.tag}</span>
                              <span className="ml-2">{item.detail}</span>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="mt-4 text-sm text-slate-600">
                          {group.items.length ? "Click the count to view animal tags." : <><span className="font-semibold text-slate-800">Nil</span><span className="ml-2">No current alert</span></>}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {/* Platform highlights */}
              <div className="rounded-[28px] border border-slate-200/70 bg-white/94 p-7 shadow-[0_14px_34px_rgba(15,23,42,0.065)]">
                <div className="flex flex-col gap-2 lg:flex-row lg:items-end lg:justify-between">
                  <div>
                    <div className="text-xs font-semibold uppercase tracking-[0.22em] text-emerald-700">Platform Highlights</div>
                    <h2 className="mt-2 text-3xl font-bold tracking-tight text-slate-950">Built for scientific herd management</h2>
                  </div>
                  <div className="max-w-xl text-sm leading-6 text-slate-600">
                    A unified platform for records, breeding, production analytics, fertility monitoring and documentation.
                  </div>
                </div>

                <div className="mt-7 grid gap-5 lg:grid-cols-2">
                  <div className="rounded-[28px] border border-emerald-100 bg-emerald-50/60 p-5 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                    <div className="flex items-start gap-5">
                      <img src="/highlight-1.png" alt="Herd Intelligence" className="ml-3 h-28 w-40 flex-shrink-0 object-contain object-center" />
                      <div>
                        <div className="text-lg font-semibold text-slate-950">1. Herd Intelligence</div>
                        <ul className="mt-3 list-disc space-y-1 pl-5 text-sm leading-7 text-slate-700">
                          <li>Centralized digital records for all animals</li>
                          <li>Real-time classification of animal</li>
                        </ul>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-[28px] border border-cyan-100 bg-cyan-50/60 p-5 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                    <div className="flex items-start gap-5">
                      <img src="/highlight-2.png" alt="Genetic and Breeding Engine" className="ml-3 h-28 w-40 flex-shrink-0 object-contain object-center" />
                      <div>
                        <div className="text-lg font-semibold text-slate-950">2. Genetic &amp; Breeding Engine</div>
                        <ul className="mt-3 list-disc space-y-1 pl-5 text-sm leading-7 text-slate-700">
                          <li>3-generation pedigree</li>
                          <li>Smart mate selection with inbreeding control</li>
                        </ul>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-[28px] border border-blue-100 bg-blue-50/60 p-5 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                    <div className="flex items-start gap-5">
                      <img src="/highlight-3.png" alt="Production Analytics" className="ml-3 h-28 w-40 flex-shrink-0 object-contain object-center" />
                      <div>
                        <div className="text-lg font-semibold text-slate-950">3. Production Analytics</div>
                        <ul className="mt-3 list-disc space-y-1 pl-5 text-sm leading-7 text-slate-700">
                          <li>Test-day based milk recording</li>
                          <li>Automated wet average &amp; herd average with parity-wise tracking</li>
                        </ul>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-[28px] border border-violet-100 bg-violet-50/60 p-5 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                    <div className="flex items-start gap-5">
                      <img src="/highlight-4.png" alt="Reproduction Monitoring" className="ml-3 h-28 w-40 flex-shrink-0 object-contain object-center" />
                      <div>
                        <div className="text-lg font-semibold text-slate-950">4. Reproduction Monitoring</div>
                        <ul className="mt-3 list-disc space-y-1 pl-5 text-sm leading-7 text-slate-700">
                          <li>AI, conception, PD tracking</li>
                          <li>Key metrics: AFC, SP, CI, fertility status</li>
                        </ul>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-[28px] border border-amber-100 bg-amber-50/60 p-5 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                    <div className="flex items-start gap-5">
                      <img src="/highlight-5.png" alt="Bull Evaluation System" className="ml-3 h-28 w-40 flex-shrink-0 object-contain object-center" />
                      <div>
                        <div className="text-lg font-semibold text-slate-950">5. Bull Evaluation System</div>
                        <ul className="mt-3 list-disc space-y-1 pl-5 text-sm leading-7 text-slate-700">
                          <li>Progeny &amp; daughter performance tracking</li>
                          <li>Current Set &amp; PT-based breeding policy</li>
                        </ul>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-[28px] border border-slate-200 bg-slate-50/80 p-5 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-lg">
                    <div className="flex items-start gap-5">
                      <img src="/highlight-6.png" alt="Animal Documentation" className="ml-3 h-28 w-40 flex-shrink-0 object-contain object-center" />
                      <div>
                        <div className="text-lg font-semibold text-slate-950">6. Animal Documentation</div>
                        <ul className="mt-3 list-disc space-y-1 pl-5 text-sm leading-7 text-slate-700">
                          <li>Linked visual identity for each animal</li>
                          <li>Automated generation of monthly and annual reports of herd performance</li>
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Footer */}
              <div className="rounded-[30px] border border-white/70 bg-slate-950 px-7 py-5 text-white shadow-[0_14px_30px_rgba(15,23,42,0.12)]">
                <div className="flex flex-col gap-2 lg:flex-row lg:items-center lg:justify-between">
                  <div>
                    <div className="text-sm font-semibold">DairyScope · Herd Intelligence Platform</div>
                    <div className="mt-1 text-xs text-slate-300">ICAR-Central Institute for Research on Buffaloes, Sub-Campus Nabha</div>
                  </div>
                  <div className="text-xs text-slate-400">Designed for precision recording, breeding decisions and institutional reporting</div>
                </div>
              </div>
            </div>
          </Section>
        )}


{sidebarView === "dashboard" && (
          <>
            <Section title="Herd Snapshot">
              <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                <table className="min-w-full text-sm">
                  <thead className="bg-slate-50/90">
                    <tr>
                      <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Category</th>
                      <th className="whitespace-nowrap px-4 py-3 text-left text-sm font-semibold text-slate-700">Count</th>
                    </tr>
                  </thead>
                  <tbody>
                    {herdSnapshotRows.map((row) => (
                      <tr key={row.label}>
                        <td className="whitespace-nowrap px-4 py-3 text-slate-700 font-medium">{row.label}</td>
                        <td className="whitespace-nowrap px-4 py-3 text-slate-700">{row.value}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>

            <Section title="Monthly Herd Strength Report">
              <div className="mb-4 grid gap-4 md:grid-cols-2">
                <SelectField label="Financial Year" value={selectedFinancialYear} onChange={setSelectedFinancialYear} options={financialYearOptions.length ? financialYearOptions : [{ value: String(new Date().getFullYear()), label: `${new Date().getFullYear()}-${String((new Date().getFullYear() + 1) % 100).padStart(2, "0")}` }]} />
                <SelectField label="Month" value={selectedHerdStructureMonth} onChange={setSelectedHerdStructureMonth} options={getFinancialMonthOptions()} />
              </div>

              <div className="space-y-5">
                <div className="rounded-[24px] border border-slate-100 bg-white/70 p-4">
                  <div className="text-sm font-semibold text-slate-700">
                    As on 1st {selectedHerdStructureMonth}, {herdStructureReport.asOfDate.getFullYear()}
                  </div>
                  <div className="mt-1 text-xs text-slate-500">Financial year-wise monthly herd strength calculated from current animals and historically present archived animals.</div>
                </div>

                <HerdStrengthTable title="Calves herd strength" rows={herdStructureReport.calfRows} totalRow={herdStructureReport.calfTotals} />

                <HerdStrengthTable title="Adults herd strength" rows={herdStructureReport.adultRows} totalRow={herdStructureReport.adultTotals} />

                <div className="rounded-[24px] border border-slate-100 bg-white/70 px-4 py-4">
                  <div className="text-xl font-bold text-slate-950">Total Herd = {herdStructureReport.calfTotals.total + herdStructureReport.adultTotals.total}</div>
                </div>
              </div>
            </Section>

            <Section title="Monthly Herd Movement Report">
              <div className="mb-4 rounded-[24px] border border-slate-100 bg-white/70 p-4">
                <div className="text-sm font-semibold text-slate-700">
                  OB + Addition (B, P, T-in) − Deduction (D, S, T-out) = CB
                </div>
                <div className="mt-1 text-xs text-slate-500">
                  OB is calculated from animals present before the selected month. CB is reconciled as OB + B + P + T-in − D − S − T-out; before month close it is a provisional balance based on entries recorded so far. Births, deaths/sales and internal transfers due to age/lactation/status change are reconciled category-wise.
                </div>
              </div>
              <HerdMovementTable report={monthlyHerdMovementReport} />
              <div className="mt-3 rounded-[24px] border border-slate-100 bg-white/70 p-4 text-xs text-slate-600">
                OB = Opening Balance; B = Birth; P = Purchase; T-in = Transfer-in; D = Death; S = Sold/Culled; T-out = Transfer-out; CB = Closing Balance. Purchased animals are counted under P using their Date of purchase / entry and age/status category on that date.
              </div>
            </Section>

            <Section title="Reproductive Status · Animals in Milk (Days in Milk)">
              <DashboardMatrixTable columns={milkStatusSummary.columns} rows={milkStatusSummary.rows} />
            </Section>

            <div className="grid gap-6 xl:grid-cols-2">
              <Section title="Reproductive Status · Dry Animals">
                <DashboardSingleTable rows={dryStatusSummary} />
              </Section>

              <Section title="Reproductive Status · Heifers (Age in Months)">
                <DashboardMatrixTable columns={heiferStatusSummary.columns} rows={heiferStatusSummary.rows} includeColostrum />
              </Section>
            </div>
          </>
        )}

        {(sidebarView === "current" || sidebarView === "archives") && (
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-[420px_1fr]">
          <Section title={registryTitle}>
            <TextField label="Search" value={search} onChange={setSearch} />
            <div className="mt-4 rounded-2xl border border-slate-100 bg-slate-50/80 px-4 py-3 text-sm text-slate-500">
              {sidebarView === "archives"
                ? `Showing ${archiveHerdSex.toLowerCase()} archived animals`
                : `Showing ${currentHerdSex.toLowerCase()} animals in the current herd`}
            </div>

            <div className="mt-4 grid gap-3">
              {registryAnimals.map((animal) => (
                <button
                  key={animal.id}
                  className={`rounded-2xl border px-4 py-4 text-left transition ${selectedId === animal.id ? "border-emerald-500 bg-emerald-50/70 shadow-sm" : "border-slate-200 bg-white/92 hover:border-emerald-300 hover:bg-emerald-50/40"}`}
                  onClick={() => { setSelectedId(animal.id); setDetailTab("pedigree"); }}
                >
                  <div className="font-semibold text-slate-900">{animal.tagNo}</div>
                  <div className="mt-1 text-sm text-slate-600">
                    {sidebarView === "archives"
                      ? `${animal.status} · ${animal.exitDate || "No exit date"}`
                      : animal.category === "Female"
                        ? `${animal.breed} · ${getFemaleLifecycle(animal)}`
                        : `${animal.breed} · ${animal.isBreedingBull === "Yes" ? `Breeding Bull (${animal.breedingSet || "Set blank"}${animal.currentSet === "No" ? (animal.isPtBull === "Yes" ? ", PT" : ", not current") : ", current"})` : "Male"}`}
                  </div>
                </button>
              ))}

              {registryAnimals.length === 0 && (
                <EmptyState
                  title={sidebarView === "archives" ? `No archived ${archiveHerdSex.toLowerCase()} animals` : `No ${currentHerdSex.toLowerCase()} animals in herd`}
                  description={sidebarView === "archives"
                    ? "Archived records matching this section will appear here."
                    : "Add animals or change the sidebar category to view a different section of the herd."}
                />
              )}
            </div>
          </Section>

          <div className="grid gap-5">
            <div className="grid gap-5 2xl:grid-cols-[minmax(0,1fr)_340px]">
              <div>
            <Section title="Selected Animal Preview">
              {uiSelectedAnimal ? (
                <div>
                  <div className="mb-3 flex justify-end gap-3">
                    <button className="rounded-xl border border-rose-300 bg-white px-4 py-2 text-rose-700 hover:bg-rose-50" onClick={deleteSelectedAnimal}>Delete Animal</button>
                    <button className="rounded-xl border border-emerald-300 bg-white px-4 py-2 text-emerald-800 hover:bg-emerald-50" onClick={startEditAnimal}>Edit Animal</button>
                  </div>
                  <div className="grid gap-4 md:grid-cols-2">
                    <div><span className="font-medium">Tag No.:</span> {uiSelectedAnimal.tagNo}</div>
                    <div><span className="font-medium">Breed:</span> {uiSelectedAnimal.breed}</div>
                    <div><span className="font-medium">DOB:</span> {uiSelectedAnimal.dob || "—"}</div>
                    <div><span className="font-medium">Sex:</span> {uiSelectedAnimal.category}</div>
                    <div><span className="font-medium">Status:</span> {uiSelectedAnimal.status}</div>
                    <div><span className="font-medium">Identification Mark:</span> {uiSelectedAnimal.identificationMark || "—"}</div>
                    {uiSelectedAnimal.category === "Female" && <div><span className="font-medium">Current female category:</span> {getFemaleLifecycle(uiSelectedAnimal)}</div>}
                    {uiSelectedAnimal.category === "Male" && <div><span className="font-medium">Breeding bull:</span> {uiSelectedAnimal.isBreedingBull === "Yes" ? `Yes (${uiSelectedAnimal.breedingSet || "Set blank"}; Current set: ${uiSelectedAnimal.currentSet || "Yes"}${uiSelectedAnimal.currentSet === "No" ? `; PT: ${uiSelectedAnimal.isPtBull || "No"}` : ""})` : "No"}</div>}
                    {uiSelectedAnimal.status !== "Active (present in herd)" && (
                      <>
                        <div><span className="font-medium">Date of Death / Culling:</span> {uiSelectedAnimal.exitDate || "—"}</div>
                        <div><span className="font-medium">Reason of Death / Culling:</span> {uiSelectedAnimal.exitReason || "—"}</div>
                      </>
                    )}
                  </div>
                </div>
              ) : (
                <div className="text-slate-500">Select an animal from the registry to unlock pedigree, reproduction, calving, production, and analytics views.</div>
              )}
            </Section>

              </div>
              {uiSelectedAnimal?.category === "Female" && <div className="2xl:sticky 2xl:top-4 2xl:self-start"><AnimalSummaryPanel animal={uiSelectedAnimal} /></div>}
            </div>

            {uiSelectedAnimal?.category === "Female" && (
              <Section title="Female Animal Tabs">
                <div className="mb-4 flex flex-wrap gap-3">
                  {FEMALE_TABS.map((tabName) => (
                    <button key={tabName} className={`rounded-xl border px-3 py-2 capitalize ${detailTab === tabName ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setDetailTab(tabName)}>
                      {tabName === "history" ? "overall history sheet" : tabName}
                    </button>
                  ))}
                </div>

                {detailTab === "photographs" && (
                  <AnimalPhotographsPanel photos={activeAnimal.femaleDetails?.photographs || []} onUpload={handleAnimalPhotoUpload} onRemove={removeAnimalPhoto} />
                )}

                {detailTab === "timeline" && (
                  <AnimalTimelinePanel animal={activeAnimal} />
                )}

                {detailTab === "pedigree" && (
                  <Grid>
                    <TextField label="Sire" value={activePedigree.sire} onChange={(v) => updateFemalePedigree("sire", v)} />
                    <TextField label="Dam" value={activePedigree.dam} onChange={(v) => updateFemalePedigree("dam", v)} />
                    <TextField label="Sire's sire" value={activePedigree.sireSire} onChange={(v) => updateFemalePedigree("sireSire", v)} />
                    <TextField label="Sire's dam" value={activePedigree.sireDam} onChange={(v) => updateFemalePedigree("sireDam", v)} />
                    <TextField label="Dam's sire" value={activePedigree.damSire} onChange={(v) => updateFemalePedigree("damSire", v)} />
                    <TextField label="Dam's dam" value={activePedigree.damDam} onChange={(v) => updateFemalePedigree("damDam", v)} />
                    <TextField label="Great-grandsire (SSS)" value={activePedigree.sireSireSire} onChange={(v) => updateFemalePedigree("sireSireSire", v)} />
                    <TextField label="Great-granddam (SSD)" value={activePedigree.sireSireDam} onChange={(v) => updateFemalePedigree("sireSireDam", v)} />
                    <TextField label="Great-grandsire (SDS)" value={activePedigree.sireDamSire} onChange={(v) => updateFemalePedigree("sireDamSire", v)} />
                    <TextField label="Great-granddam (SDD)" value={activePedigree.sireDamDam} onChange={(v) => updateFemalePedigree("sireDamDam", v)} />
                    <TextField label="Great-grandsire (DSS)" value={activePedigree.damSireSire} onChange={(v) => updateFemalePedigree("damSireSire", v)} />
                    <TextField label="Great-granddam (DSD)" value={activePedigree.damSireDam} onChange={(v) => updateFemalePedigree("damSireDam", v)} />
                    <TextField label="Great-grandsire (DDS)" value={activePedigree.damDamSire} onChange={(v) => updateFemalePedigree("damDamSire", v)} />
                    <TextField label="Great-granddam (DDD)" value={activePedigree.damDamDam} onChange={(v) => updateFemalePedigree("damDamDam", v)} />
                  </Grid>
                )}

                {detailTab === "reproduction" && selectedReproParity && reproSummary && (
                  <div className="space-y-5">
                    {[...(activeAnimal.femaleDetails.reproductionParities || [])]
                      .filter((rp) => shouldShowReproParity(activeAnimal, rp))
                      .sort((a, b) => Number(a.parityNo || 0) - Number(b.parityNo || 0))
                      .map((rp) => {
                        const reproClosed = isReproParityClosedByProductiveCalving(activeAnimal, rp.parityNo);
                        const reproKey = `${activeAnimal.id}:repro:${rp.parityNo}`;
                        const isSelected = String(activeAnimal.femaleDetails.selectedReproParity) === String(rp.parityNo);
                        const rpAI = sortAIRecords(rp.aiRecords || []);
                        const rpLastAI = rpAI[rpAI.length - 1]?.aiDate || "—";
                        const rpConception = getConceptionDateForReproParity(rp) || "—";
                        const isMinimized = Boolean(!isSelected || (reproClosed && expandedClosedRepro[reproKey] === false && !closedParityEdit[reproKey]));

                        if (isMinimized) {
                          return (
                            <div key={`rp-mini-${rp.parityNo}`} className={`group relative overflow-hidden rounded-[30px] border bg-gradient-to-br p-5 shadow-[0_18px_42px_rgba(15,23,42,0.08)] transition duration-300 hover:-translate-y-0.5 hover:shadow-[0_22px_50px_rgba(15,23,42,0.1)] ${reproClosed ? "border-amber-200 from-amber-50 via-white to-emerald-50/50" : "border-emerald-200 from-white via-emerald-50/70 to-cyan-50/40"}`}>
                              <div className="absolute -right-10 -top-10 h-24 w-24 rounded-full bg-emerald-100/60 blur-2xl" />
                              <div className="relative z-10 flex flex-wrap items-start justify-between gap-3">
                                <div className="space-y-2">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <span className="rounded-full bg-white/95 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.2em] text-emerald-700 shadow-sm">Parity {rp.parityNo}</span>
                                    <span className={`rounded-full px-3 py-1 text-xs font-semibold ${reproClosed ? "bg-amber-100 text-amber-800" : "bg-emerald-100 text-emerald-800"}`}>{reproClosed ? "Closed" : "Open"}</span>
                                  </div>
                                  <div className="text-xl font-bold tracking-tight text-emerald-950">Repro parity {rp.parityNo}</div>
                                  <div className="text-sm leading-6 text-slate-600">Last AI: <span className="font-semibold text-slate-800">{rpLastAI}</span> · Conception: <span className="font-semibold text-slate-800">{rpConception}</span> · {reproClosed ? "closed by productive calving" : "ready for next insemination"}</div>
                                </div>
                                <button title="Expand parity" aria-label="Expand parity" className={iconButtonClass("emerald", "lg")} onClick={() => { setExpandedClosedRepro((prev) => ({ ...prev, [reproKey]: true })); selectReproParity(rp.parityNo); }}><span aria-hidden="true">⌄</span></button>
                              </div>
                            </div>
                          );
                        }

                        return (
                          <div key={`rp-expanded-${rp.parityNo}`} className={`overflow-hidden rounded-[30px] border bg-gradient-to-br p-5 shadow-[0_18px_42px_rgba(15,23,42,0.08)] ${reproClosed ? "border-amber-200 from-amber-50/80 via-white to-emerald-50/30" : "border-emerald-200 from-white via-emerald-50/70 to-cyan-50/40"}`}>
                            <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
                              <div className="flex flex-wrap items-center gap-3">
                                <span className="rounded-full bg-white/95 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.2em] text-emerald-700 shadow-sm">Parity {rp.parityNo}</span>
                                <div className="text-xl font-bold tracking-tight text-emerald-950">Repro parity {rp.parityNo}</div>
                                <span className={`rounded-full px-3 py-1 text-xs font-semibold ${reproClosed ? "bg-amber-100 text-amber-800" : "bg-emerald-100 text-emerald-800"}`}>{reproClosed ? "Closed" : "Open"}</span>
                              </div>
                              <div className="flex gap-2">
                                {reproClosed && (
                                  <button
                                    className={iconButtonClass("emerald", "sm")}
                                    title="Minimise parity"
                                    aria-label="Minimise parity"
                                    onClick={() => setExpandedClosedRepro((prev) => ({ ...prev, [reproKey]: false }))}
                                  >
                                    <span aria-hidden="true">⌃</span>
                                  </button>
                                )}
                                {reproClosed && (
                                  <button
                                    className={iconButtonClass("amber", "sm")}
                                    title={selectedReproParityLocked ? "Edit closed parity" : "Lock parity"}
                                    aria-label={selectedReproParityLocked ? "Edit closed parity" : "Lock parity"}
                                    onClick={() => setClosedParityEdit((prev) => ({ ...prev, [reproKey]: !prev[reproKey] }))}
                                  >
                                    <span aria-hidden="true">{selectedReproParityLocked ? "✎" : "🔒"}</span>
                                  </button>
                                )}
                              </div>
                            </div>

                            {reproClosed && (
                              <div className="mb-4 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">
                                <span className="font-semibold">Repro parity closed.</span>
                                <span className="ml-2">Closed by productive calving parity {Number(rp.parityNo) + 1}. The next repro parity has been opened automatically.</span>
                              </div>
                            )}

                            <div className="rounded-[24px] border border-emerald-100 bg-white/80 p-4 shadow-sm">
                              <div className="grid gap-4 md:grid-cols-3">
                                <TextField label="Repro parity" value={String(rp.parityNo || "")} onChange={() => {}} readOnly />
                                <TextField label="Stage" value={reproSummary.parityLabel} onChange={() => {}} readOnly />
                                <TextField label="Last AI date" value={reproSummary.lastAIDate || ""} onChange={() => {}} readOnly />
                                <DateField label="Conception date" value={selectedReproParity.conceptionDate || ""} onChange={() => {}} readOnly />
                                <TextField label="Expected calving date" value={selectedReproParity.expectedCalvingDate || ""} onChange={() => {}} readOnly />
                                <TextField label="Number of services" value={String(reproSummary.services)} onChange={() => {}} readOnly />
                              </div>
                              <div className="mt-3"><AreaField label="Remarks" value={selectedReproParity.remarks || ""} onChange={(v) => updateSelectedRepro("remarks", v)} rows={3} readOnly={selectedReproParityLocked} /></div>
                            </div>

                            <div className="mt-5 rounded-[24px] border border-emerald-100 bg-white p-4 shadow-sm">
                              <div className="mb-3 text-base font-semibold text-emerald-900">Mate Selection</div>
                              <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                                <table className="min-w-full border-collapse text-sm">
                                  <thead className="bg-slate-50/90">
                                    <tr className="bg-emerald-50 text-left text-emerald-900">
                                      <th className="border border-emerald-100 px-2 py-2">Rank</th>
                                      <th className="border border-emerald-100 px-2 py-2">Bull No.</th>
                                      <th className="border border-emerald-100 px-2 py-2">Set No.</th>
                                      <th className="border border-emerald-100 px-2 py-2">Estimated Inbreeding %</th>
                                      <th className="border border-emerald-100 px-2 py-2">Status</th>
                                      <th className="border border-emerald-100 px-2 py-2">Recommendation</th>
                                      <th className="border border-emerald-100 px-2 py-2">Reason</th>
                                      <th className="border border-emerald-100 px-2 py-2">Action</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {mateSelectionRows.map((row) => (
                                      <tr key={row.bull.id}>
                                        <td className="border border-emerald-100 px-2 py-2">{row.autoRank}</td>
                                        <td className="border border-emerald-100 px-2 py-2">{row.bull.tagNo}</td>
                                        <td className="border border-emerald-100 px-2 py-2">{row.bull.breedingSet || "—"}</td>
                                        <td className="border border-emerald-100 px-2 py-2">{row.inbreedingPct}</td>
                                        <td className="border border-emerald-100 px-2 py-2">{row.status}</td>
                                        <td className="border border-emerald-100 px-2 py-2">{row.recommendation}</td>
                                        <td className="border border-emerald-100 px-2 py-2">{row.reason}</td>
                                        <td className="border border-emerald-100 px-2 py-2">
                                          {row.status !== "Avoid" ? (
                                            <button
                                              className="rounded-xl border border-emerald-300 bg-white px-3 py-1 text-emerald-800 hover:bg-emerald-50 disabled:opacity-40"
                                              onClick={() => applyBullToSelectedParity(row.bull)}
                                              disabled={selectedReproParityLocked}
                                            >
                                              Use
                                            </button>
                                          ) : (
                                            <span className="text-slate-400">Blocked</span>
                                          )}
                                        </td>
                                      </tr>
                                    ))}
                                    {mateSelectionRows.length === 0 && (
                                      <tr>
                                        <td colSpan={8} className="border border-emerald-100 px-2 py-3 text-center text-slate-500">
                                          No eligible current-set or PT bulls available for this female.
                                        </td>
                                      </tr>
                                    )}
                                  </tbody>
                                </table>
                              </div>
                            </div>

                            <div className="mt-5">
                              <div className="mb-2 font-semibold">AI details in selected parity</div>
                              <div className="grid gap-4">
                                {selectedReproParity.aiRecords.length === 0 ? <div className="rounded-xl border border-emerald-100 bg-white p-3 text-sm text-slate-500">No AI attempts entered yet for this parity.</div> : selectedReproParity.aiRecords.map((rec, idx) => (
                                  <div key={`ai-${idx}`} className="rounded-2xl border border-emerald-100 bg-white p-3">
                                    <Grid>
                                      <DateField label={`AI ${idx + 1} date`} value={rec.aiDate || ""} onChange={(v) => updateAIRecord(idx, "aiDate", v)} readOnly={selectedReproParityLocked} />
                                      <TextField label="Bull No." value={rec.aiBullNo || ""} onChange={(v) => updateAIRecord(idx, "aiBullNo", v)} readOnly={selectedReproParityLocked} />
                                      <TextField label="Set No." value={rec.aiSetNo || ""} onChange={(v) => updateAIRecord(idx, "aiSetNo", normalizeRomanInput(v))} readOnly={selectedReproParityLocked} />
                                      <SelectField label="Result" value={rec.result || "Pending"} onChange={(v) => updateAIRecord(idx, "result", v)} options={AI_RESULTS} disabled={selectedReproParityLocked} />
                                    </Grid>
                                  </div>
                                ))}
                              </div>
                              <div className="mt-3 flex gap-3">
                                <button className="flex h-14 w-14 items-center justify-center rounded-2xl bg-emerald-600 text-4xl font-bold text-white hover:bg-emerald-700 disabled:opacity-40" onClick={addAIRecord} disabled={selectedReproParityLocked}>+</button>
                                <button className="flex h-14 w-14 items-center justify-center rounded-2xl border border-slate-300 bg-white text-4xl font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40" onClick={removeAIRecord} disabled={selectedReproParityLocked}>−</button>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                  </div>
                )}

                {detailTab === "calving" && (
                  <div className="space-y-5">
                    {activeAnimal.femaleDetails.calvingParities.map((cp, idx) => {
                      const parityEvents = getCalvingOutcomeEvents(cp);
                      const productiveClosure = parityEvents.find((event) => isCompleteProductiveCalvingEvent(event));
                      const readyForClosure = parityEvents.find((event) => !event.finalized && isReadyToFinalizeProductiveCalvingEvent(event));
                      const calvingEditKey = `${activeAnimal.id}:calving:${cp.parityNo}`;
                      const calvingClosed = Boolean(productiveClosure);
                      const calvingLocked = Boolean(calvingClosed && !closedParityEdit[calvingEditKey]);
                      const calvingExpanded = Boolean(expandedClosedCalving[calvingEditKey] || closedParityEdit[calvingEditKey]);
                      const showCollapsedClosed = Boolean(calvingClosed && !calvingExpanded);
                      const productiveSummary = productiveClosure || {};
                      return (
                        <div key={`calving-${cp.parityNo}`} className={`overflow-hidden rounded-[30px] border bg-gradient-to-br p-5 shadow-[0_18px_42px_rgba(15,23,42,0.08)] ${calvingClosed ? "border-amber-200 from-amber-50/80 via-white to-emerald-50/30" : "border-emerald-200 from-white via-emerald-50/60 to-cyan-50/30"}`}>
                          <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                            <div className="flex items-center gap-3"><span className="rounded-full bg-white/90 px-3 py-1 text-xs font-bold uppercase tracking-[0.18em] text-emerald-700 shadow-sm">Parity {cp.parityNo}</span><div className="text-lg font-bold tracking-tight text-emerald-950">Calving parity {cp.parityNo}</div></div>
                            {readyForClosure && !calvingClosed && (
                              <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-800">Ready to close on Next Tab</span>
                            )}
                            {calvingClosed && (
                              <div className="flex items-center gap-2">
                                <span className="rounded-full bg-amber-50 px-3 py-1 text-xs font-semibold text-amber-800">Closed productive calving</span>
                                <button
                                  className={iconButtonClass("emerald", "sm")}
                                  onClick={() => setExpandedClosedCalving((prev) => ({ ...prev, [calvingEditKey]: !prev[calvingEditKey] }))}
                                >
                                  <span aria-hidden="true">{showCollapsedClosed ? "⌄" : "⌃"}</span>
                                </button>
                                {!showCollapsedClosed && (
                                  <button
                                    className={iconButtonClass("amber", "sm")}
                                    onClick={() => setClosedParityEdit((prev) => ({ ...prev, [calvingEditKey]: !prev[calvingEditKey] }))}
                                  >
                                    <span aria-hidden="true">{calvingLocked ? "✎" : "🔒"}</span>
                                  </button>
                                )}
                              </div>
                            )}
                          </div>

                          {showCollapsedClosed ? (
                            <div className="grid gap-3 rounded-[24px] border border-amber-200 bg-gradient-to-r from-amber-50 via-white to-amber-50/70 p-5 text-sm shadow-[inset_0_1px_0_rgba(255,255,255,0.8)] md:grid-cols-4">
                              <div><span className="font-semibold text-slate-500">Outcome</span><div className="font-bold text-slate-900">{productiveSummary.calvingOutcome || "—"}</div></div>
                              <div><span className="font-semibold text-slate-500">Date</span><div className="font-bold text-slate-900">{productiveSummary.calvingDate || "—"}</div></div>
                              <div><span className="font-semibold text-slate-500">Calf</span><div className="font-bold text-slate-900">{isCalfIdentityOutcome(productiveSummary) ? `${productiveSummary.calfSex || "—"} ${productiveSummary.calfTag ? `· ${productiveSummary.calfTag}` : ""}` : "—"}</div></div>
                              <div><span className="font-semibold text-slate-500">Status</span><div className="font-bold text-amber-800">Locked</div></div>
                            </div>
                          ) : (
                          <div className="space-y-4">
                            {parityEvents.map((event, outcomeIdx) => {
                              const metrics = computeCalvingOutcomeMetrics(selectedAnimal, cp.parityNo, event, outcomeIdx);
                              const autoCalfSire = getCalfSireForCalving(selectedAnimal, cp.parityNo, outcomeIdx);
                              const showCalfIdentityFields = isCalfIdentityOutcome(event);
                              const afterAbortion = outcomeIdx > 0 && parityEvents.slice(0, outcomeIdx).some((prev) => prev.calvingOutcome === "Abortion");
                              return (
                                <div key={`calving-${cp.parityNo}-outcome-${outcomeIdx}`} className={`rounded-[24px] border p-5 shadow-sm ${afterAbortion ? "border-rose-200 bg-gradient-to-br from-rose-50/80 to-white" : "border-emerald-100 bg-gradient-to-br from-white to-emerald-50/40"}`}>
                                  <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
                                    <div className="font-semibold text-slate-900">Outcome / Conception {outcomeIdx + 1}</div>
                                    {afterAbortion && <span className="rounded-full bg-rose-100 px-3 py-1 text-xs font-semibold text-rose-700">After previous abortion</span>}
                                  </div>
                                  <Grid>
                                    <DateField label="Calving/outcome date" value={event.calvingDate || ""} onChange={(v) => updateCalvingOutcomeEvent(idx, outcomeIdx, "calvingDate", v)} readOnly={calvingLocked} />
                                    <SelectField label="Outcome" value={event.calvingOutcome || ""} onChange={(v) => updateCalvingOutcomeEvent(idx, outcomeIdx, "calvingOutcome", v)} options={CALVING_OUTCOMES} disabled={calvingLocked} />
                                    <TextField label="AFC (days)" value={metrics.afc ?? ""} onChange={() => {}} readOnly />
                                    <TextField label="Gestation period (days)" value={metrics.gestationPeriod ?? ""} onChange={() => {}} readOnly />
                                    <TextField label="Service period (days)" value={metrics.servicePeriod ?? ""} onChange={() => {}} readOnly />
                                    <TextField label="Calving interval (days)" value={metrics.calvingInterval ?? ""} onChange={() => {}} readOnly />
                                    {showCalfIdentityFields && (
                                      <>
                                        <SelectField label="Calf sex" value={event.calfSex || ""} onChange={(v) => updateCalvingOutcomeEvent(idx, outcomeIdx, "calfSex", v)} options={["", ...SEX_OPTIONS]} disabled={calvingLocked} />
                                        <TextField label="Calf tag no. (auto-adds calf to herd)" value={event.calfTag || ""} onChange={(v) => updateCalvingOutcomeEvent(idx, outcomeIdx, "calfTag", v)} readOnly={calvingLocked} />
                                        <TextField label="Calf sire (auto)" value={autoCalfSire || event.calfSire || ""} onChange={() => {}} readOnly />
                                      </>
                                    )}
                                  </Grid>
                                  <div className="mt-3"><AreaField label="Remarks" value={event.remarks || ""} onChange={(v) => updateCalvingOutcomeEvent(idx, outcomeIdx, "remarks", v)} rows={3} readOnly={calvingLocked} /></div>
                                </div>
                              );
                            })}
                          </div>
                          )}

                          {!calvingLocked && parityEvents.some((event) => event.calvingOutcome === "Abortion" && event.calvingDate) && (
                            <div className="mt-3 flex gap-3">
                              <button className="rounded-2xl bg-emerald-600 px-4 py-3 text-sm font-bold text-white hover:bg-emerald-700" onClick={() => addCalvingOutcomeEvent(idx)}>+ Add outcome after abortion</button>
                              {parityEvents.length > 1 && <button className="rounded-2xl border border-slate-300 bg-white px-4 py-3 text-sm font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeCalvingOutcomeEvent(idx)}>− Remove outcome</button>}
                            </div>
                          )}
                        </div>
                      );
                    })}
                    <div className="rounded-2xl border border-emerald-100 bg-emerald-50/70 p-4 text-sm text-emerald-900">
                      Next calving parity will appear automatically after insemination/conception details are entered in the current reproductive parity and <span className="font-semibold">Next Tab</span> is clicked.
                    </div>
                  </div>
                )}

                {detailTab === "production" && !selectedLactation && (
                  <EmptyState title="No production parity yet" description="Production/lactation parity starts only after Normal calving, Stillbirth or Premature calving. Abortion closes pregnancy but does not start production." />
                )}

                {detailTab === "production" && selectedLactation && (
                  <div className="space-y-5">
                    {[...(activeAnimal.femaleDetails.productionLactations || [])]
                      .sort((a, b) => Number(a.parityNo || 0) - Number(b.parityNo || 0))
                      .map((lac) => {
                        const key = `${activeAnimal.id}:production:${lac.parityNo}`;
                        const isExpanded = String(activeAnimal.femaleDetails.selectedProductionParity) === String(lac.parityNo) && expandedProductionParity[key] !== false;
                        const metrics = computeProductionMetrics(lac);
                        const tlmySummary = lac.entryMode === "Manual"
                          ? (lac.manualSummary?.totalLactationMilk || "—")
                          : metrics.totalLactationMilk.toFixed(1);
                        const slmySummary = lac.entryMode === "Manual"
                          ? (lac.manualSummary?.standardLactationMilk || "—")
                          : metrics.standardLactationMilk.toFixed(1);

                        if (!isExpanded) {
                          return (
                            <div key={`prod-mini-${lac.parityNo}`} className="group relative overflow-hidden rounded-[30px] border border-amber-200 bg-gradient-to-br from-amber-50 via-white to-emerald-50/50 p-5 shadow-[0_18px_42px_rgba(15,23,42,0.08)] transition duration-300 hover:-translate-y-0.5 hover:shadow-[0_22px_50px_rgba(15,23,42,0.1)]">
                              <div className="absolute -right-10 -top-10 h-28 w-28 rounded-full bg-amber-100/70 blur-2xl" />
                              <div className="absolute -bottom-12 left-1/3 h-24 w-24 rounded-full bg-emerald-100/60 blur-2xl" />
                              <div className="relative z-10 flex flex-wrap items-start justify-between gap-4">
                                <div className="space-y-3">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <span className="rounded-full bg-white/95 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.2em] text-emerald-700 shadow-sm">Parity {lac.parityNo}</span>
                                    <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-800">{lac.entryMode || "Manual"}</span>
                                    {lac.dryDate ? <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">Dry completed</span> : <span className="rounded-full bg-cyan-50 px-3 py-1 text-xs font-semibold text-cyan-800">Open lactation</span>}
                                  </div>
                                  <div>
                                    <div className="text-xl font-bold tracking-tight text-emerald-950">Production parity {lac.parityNo}</div>
                                    <div className="mt-1 text-sm leading-6 text-slate-600">Neatly collapsed summary card with instant glance at calving, dry status and milk totals.</div>
                                  </div>
                                  <div className="flex flex-wrap gap-2 text-sm">
                                    <span className="rounded-2xl border border-emerald-100 bg-white/85 px-3 py-2 font-medium text-slate-700">Calving: <span className="font-bold text-slate-900">{lac.calvingDate || "—"}</span></span>
                                    <span className="rounded-2xl border border-emerald-100 bg-white/85 px-3 py-2 font-medium text-slate-700">Dry: <span className="font-bold text-slate-900">{lac.dryDate || "—"}</span></span>
                                    <span className="rounded-2xl border border-emerald-100 bg-white/85 px-3 py-2 font-medium text-slate-700">TLMY: <span className="font-bold text-emerald-900">{tlmySummary}</span></span>
                                    <span className="rounded-2xl border border-emerald-100 bg-white/85 px-3 py-2 font-medium text-slate-700">SLMY: <span className="font-bold text-emerald-900">{slmySummary}</span></span>
                                  </div>
                                </div>
                                <button
                                  className={iconButtonClass("emerald", "lg")}
                                  onClick={() => {
                                    setExpandedProductionParity((prev) => ({ ...prev, [key]: true }));
                                    selectProductionParity(lac.parityNo);
                                  }}
                                 title="Expand parity" aria-label="Expand parity">
                                  <span aria-hidden="true">⌄</span>
                                </button>
                              </div>
                            </div>
                          );
                        }

                        return (
                          <div key={`prod-expanded-${lac.parityNo}`} className="group relative overflow-hidden rounded-[30px] border border-emerald-200 bg-gradient-to-br from-white via-emerald-50/60 to-cyan-50/35 p-[1px] shadow-[0_20px_52px_rgba(15,23,42,0.1)]">
                            <div className="rounded-[29px] bg-white/95 p-5 space-y-5">
                              <div className="relative overflow-hidden rounded-[26px] border border-emerald-100 bg-gradient-to-r from-emerald-50 via-white to-cyan-50/70 p-5">
                                <div className="absolute -right-6 -top-8 h-24 w-24 rounded-full bg-emerald-100/70 blur-2xl" />
                                <div className="relative z-10 flex flex-wrap items-start justify-between gap-4">
                                  <div className="space-y-3">
                                    <div className="flex flex-wrap items-center gap-2">
                                      <span className="rounded-full bg-white/95 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.2em] text-emerald-700 shadow-sm">Parity {lac.parityNo}</span>
                                      <span className="rounded-full bg-emerald-100 px-3 py-1 text-xs font-semibold text-emerald-800">Expanded</span>
                                      <span className="rounded-full bg-white/95 px-3 py-1 text-xs font-semibold text-slate-700 shadow-sm">{lac.entryMode || "Manual"}</span>
                                    </div>
                                    <div>
                                      <div className="text-2xl font-bold tracking-tight text-emerald-950">Production parity {lac.parityNo}</div>
                                      <div className="mt-1 text-sm leading-6 text-slate-600">Detailed lactation panel with editable production inputs and computed parity-wise milk summary.</div>
                                    </div>
                                    <div className="flex flex-wrap gap-2 text-sm">
                                      <span className="rounded-2xl border border-emerald-100 bg-white/90 px-3 py-2 font-medium text-slate-700">Calving: <span className="font-bold text-slate-900">{lac.calvingDate || "—"}</span></span>
                                      <span className="rounded-2xl border border-emerald-100 bg-white/90 px-3 py-2 font-medium text-slate-700">Dry: <span className="font-bold text-slate-900">{lac.dryDate || "—"}</span></span>
                                      <span className="rounded-2xl border border-emerald-100 bg-white/90 px-3 py-2 font-medium text-slate-700">Current TLMY: <span className="font-bold text-emerald-900">{tlmySummary}</span></span>
                                    </div>
                                  </div>
                                  <button
                                    className={iconButtonClass("emerald", "lg")}
                                    onClick={() => setExpandedProductionParity((prev) => ({ ...prev, [key]: false }))}
                                   title="Minimise parity" aria-label="Minimise parity">
                                    <span aria-hidden="true">⌃</span>
                                  </button>
                                </div>
                              </div>

                              <div className="rounded-[24px] border border-emerald-100 bg-emerald-50/70 p-4 shadow-sm">
                                <Grid>
                                  <TextField label="Parity no." value={String(lac.parityNo || "")} onChange={() => {}} readOnly />
                                  <DateField label="Calving date" value={lac.calvingDate || ""} onChange={() => {}} readOnly />
                                  <DateField label="Dry date" value={lac.dryDate || ""} onChange={(v) => updateSelectedLactation("dryDate", v)} />
                                  <SelectField label="Entry mode" value={lac.entryMode || "Manual"} onChange={(v) => updateSelectedLactation("entryMode", v)} options={ENTRY_MODES} />
                                  <TextField label="Lactation length (days)" value={metrics.lactationLength || ""} onChange={() => {}} readOnly />
                                  <TextField label="Peak yield" value={metrics.peakYield || ""} onChange={() => {}} readOnly />
                                </Grid>
                              </div>

                              {lac.entryMode === "Manual" ? (
                                <div className="rounded-[24px] border border-emerald-100 bg-white p-4 shadow-sm">
                                  <Grid>
                                    <TextField label="Total lactation milk" value={lac.manualSummary.totalLactationMilk || ""} onChange={(v) => updateManualSummary("totalLactationMilk", v)} />
                                    <TextField label="Standard lactation milk" value={lac.manualSummary.standardLactationMilk || ""} onChange={(v) => updateManualSummary("standardLactationMilk", v)} />
                                    <TextField label="Peak yield" value={lac.manualSummary.peakYield || ""} onChange={(v) => updateManualSummary("peakYield", v)} />
                                  </Grid>
                                </div>
                              ) : (
                                <div className="space-y-3">
                                  <div className="rounded-[24px] border border-emerald-100 bg-white p-4 shadow-sm">
                                    <div className="mb-3 text-sm text-slate-600">First recordable Friday: {firstRecordableFriday(lac.calvingDate) || "—"}</div>
                                    <div className="grid gap-4">
                                      {(lac.fridayRecords || []).length === 0 ? (
                                        <div className="rounded-xl border border-emerald-100 bg-white p-3 text-sm text-slate-500">Press + to generate the first recordable Friday and successive Friday rows.</div>
                                      ) : lac.fridayRecords.map((rec, idx) => (
                                        <div key={`fr-${idx}`} className="rounded-[22px] border border-emerald-100 bg-emerald-50/70 p-3">
                                          <Grid>
                                            <TextField label={`Friday ${idx + 1} date`} value={rec.date || ""} onChange={() => {}} readOnly />
                                            <TextField label="Morning milk" value={rec.morningMilk || ""} onChange={(v) => updateFridayRecord(idx, "morningMilk", v)} />
                                            <TextField label="Evening milk" value={rec.eveningMilk || ""} onChange={(v) => updateFridayRecord(idx, "eveningMilk", v)} />
                                            <TextField label="Total MY" value={rec.totalDailyYield || ""} onChange={() => {}} readOnly />
                                          </Grid>
                                        </div>
                                      ))}
                                    </div>
                                    <div className="mt-3 flex gap-3">
                                      <button className="flex h-14 w-14 items-center justify-center rounded-2xl bg-emerald-600 text-4xl font-bold text-white shadow-sm hover:bg-emerald-700" onClick={addFridayRecord}>+</button>
                                      <button className="flex h-14 w-14 items-center justify-center rounded-2xl border border-slate-300 bg-white text-4xl font-bold text-slate-700 shadow-sm hover:bg-slate-50" onClick={removeFridayRecord}>−</button>
                                    </div>
                                  </div>
                                </div>
                              )}

                              <div className="rounded-[24px] border border-emerald-100 bg-gradient-to-br from-emerald-50/80 to-white p-4 shadow-sm">
                                <div className="mb-3 flex items-center justify-between gap-3">
                                  <div>
                                    <div className="text-sm font-semibold uppercase tracking-[0.18em] text-emerald-700">Computed production summary</div>
                                    <div className="mt-1 text-sm text-slate-600">Auto-computed parity-level milk figures for fast review.</div>
                                  </div>
                                  <button className="rounded-xl border border-emerald-300 bg-white px-4 py-2 text-sm font-semibold text-emerald-800 hover:bg-emerald-50" onClick={() => setShowMonthlyTLMBreakup((prev) => !prev)}>Month-wise split of TLMY</button>
                                </div>
                                <Grid>
                                  <TextField label="TLMY" value={lac.entryMode === "Manual" ? (lac.manualSummary.totalLactationMilk || "") : metrics.totalLactationMilk.toFixed(1)} onChange={() => {}} readOnly />
                                  <TextField label="SLMY" value={lac.entryMode === "Manual" ? (lac.manualSummary.standardLactationMilk || "") : metrics.standardLactationMilk.toFixed(1)} onChange={() => {}} readOnly />
                                  <TextField label="Peak yield" value={lac.entryMode === "Manual" ? (lac.manualSummary.peakYield || "") : metrics.peakYield.toFixed(1)} onChange={() => {}} readOnly />
                                </Grid>
                                {showMonthlyTLMBreakup && (
                                  <div className="mt-3 overflow-x-auto rounded-2xl border border-emerald-100 bg-white">
                                    {monthlyTLMBreakup.length ? (
                                      <table className="min-w-full border-collapse text-sm">
                                        <thead className="bg-emerald-50 text-emerald-900">
                                          <tr>
                                            <th className="border border-emerald-100 px-3 py-2 text-left">Month</th>
                                            <th className="border border-emerald-100 px-3 py-2 text-left">Milk yield</th>
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {monthlyTLMBreakup.map((row) => (
                                            <tr key={row.month}>
                                              <td className="border border-emerald-100 px-3 py-2">{row.month}</td>
                                              <td className="border border-emerald-100 px-3 py-2">{row.milkYield.toFixed(1)}</td>
                                            </tr>
                                          ))}
                                        </tbody>
                                      </table>
                                    ) : (
                                      <div className="rounded-xl border border-emerald-100 bg-emerald-50 p-3 text-sm text-slate-600">No month-wise split is available yet. Add Friday records and dry date as needed.</div>
                                    )}
                                  </div>
                                )}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                  </div>
                )}


                {detailTab === "health" && selectedHealth && (
                  <div className="space-y-5">
                    <div className="flex flex-wrap gap-3">
                      {HEALTH_SUBTABS.map((sub) => <button key={sub.id} className={`rounded-xl border px-3 py-2 ${healthSubTab === sub.id ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setHealthSubTab(sub.id)}>{sub.label}</button>)}
                    </div>
                    {healthSubTab === "bodyWeight" && (
                      <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Body weight and recording date</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(femaleTarget, "bodyWeightRecords", makeBodyWeightRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(femaleTarget, "bodyWeightRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.bodyWeightRecords.map((rec, idx) => <div key={`bw-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Recording date" value={rec.recordDate || ""} onChange={(v) => updateHealthRecord(femaleTarget, "bodyWeightRecords", idx, "recordDate", v)} /><TextField label="Body weight" value={rec.bodyWeight || ""} onChange={(v) => updateHealthRecord(femaleTarget, "bodyWeightRecords", idx, "bodyWeight", v)} /></Grid></div>)}</div></div>
                    )}
                    {healthSubTab === "deworming" && (
                      <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Deworming date and anthelmintic used</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(femaleTarget, "dewormingRecords", makeDewormingRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(femaleTarget, "dewormingRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.dewormingRecords.map((rec, idx) => <div key={`dw-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Deworming date" value={rec.dewormingDate || ""} onChange={(v) => updateHealthRecord(femaleTarget, "dewormingRecords", idx, "dewormingDate", v)} /><TextField label="Anthelmintic used" value={rec.anthelminticUsed || ""} onChange={(v) => updateHealthRecord(femaleTarget, "dewormingRecords", idx, "anthelminticUsed", v)} /></Grid></div>)}</div></div>
                    )}
                    {healthSubTab === "vaccination" && (
                      <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Vaccination date and vaccine used</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(femaleTarget, "vaccinationRecords", makeVaccinationRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(femaleTarget, "vaccinationRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.vaccinationRecords.map((rec, idx) => <div key={`vac-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Vaccination date" value={rec.vaccinationDate || ""} onChange={(v) => updateHealthRecord(femaleTarget, "vaccinationRecords", idx, "vaccinationDate", v)} /><TextField label="Vaccine used" value={rec.vaccineUsed || ""} onChange={(v) => updateHealthRecord(femaleTarget, "vaccinationRecords", idx, "vaccineUsed", v)} /></Grid></div>)}</div></div>
                    )}
                    {healthSubTab === "treatment" && (
                      <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Treatment dates, diagnosis and treatment given</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(femaleTarget, "treatmentRecords", makeTreatmentRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(femaleTarget, "treatmentRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.treatmentRecords.map((rec, idx) => <div key={`tx-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Treatment date" value={rec.treatmentDate || ""} onChange={(v) => updateHealthRecord(femaleTarget, "treatmentRecords", idx, "treatmentDate", v)} /><TextField label="Diagnosis" value={rec.diagnosis || ""} onChange={(v) => updateHealthRecord(femaleTarget, "treatmentRecords", idx, "diagnosis", v)} /><TextField label="Treatment given" value={rec.treatmentGiven || ""} onChange={(v) => updateHealthRecord(femaleTarget, "treatmentRecords", idx, "treatmentGiven", v)} /></Grid></div>)}</div></div>
                    )}
                  </div>
                )}

                {detailTab === "history" && (
                  <div className="space-y-5">
                    <div className="flex justify-end">
                      <button
                        className="rounded-xl border border-emerald-300 bg-white px-4 py-2 text-sm font-medium text-emerald-800 hover:bg-emerald-50"
                        onClick={() => exportHistoryPdf(activeAnimal)}
                      >
                        Export History PDF
                      </button>
                    </div>
                    <Grid>
                      <TextField label="AFC (days)" value={String(afcValue || "")} onChange={() => {}} readOnly />
                      <TextField label="Reason for culling" value={activeAnimal.femaleDetails.historyMeta.reasonForCulling || ""} onChange={(v) => updateFemaleHistoryMeta("reasonForCulling", v)} />
                      <TextField label="Book value" value={activeAnimal.femaleDetails.historyMeta.bookValue || ""} onChange={(v) => updateFemaleHistoryMeta("bookValue", v)} />
                    </Grid>
                    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3"><table className="min-w-full border-collapse text-sm"><thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Parity","Date of 1st AI","No. of AI till Conception","Date of Conception","Bull No./Set No. leading to Conception","Conception No.","Outcome Type","Outcome Date","GP","Sex of Calf","Tag No. of Calf","Dry Date","Dry Period","TLMY","SLMY","LL","PY","SP","CI"].map((h) => <th key={h} className="whitespace-nowrap border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead><tbody>{historyRows.map((row, idx) => <tr key={`hist-${idx}`} className="hover:bg-emerald-50/50"><td className="border border-emerald-100 px-2 py-2">{row.parity}</td><td className="border border-emerald-100 px-2 py-2">{row.firstAI}</td><td className="border border-emerald-100 px-2 py-2">{row.aiTillConception}</td><td className="border border-emerald-100 px-2 py-2">{row.conceptionDate}</td><td className="border border-emerald-100 px-2 py-2">{row.bullNo}</td><td className="border border-emerald-100 px-2 py-2">{row.totalAI}</td><td className="border border-emerald-100 px-2 py-2">{row.outcomeType}</td><td className="border border-emerald-100 px-2 py-2">{row.dateCalved}</td><td className="border border-emerald-100 px-2 py-2">{row.gp}</td><td className="border border-emerald-100 px-2 py-2">{row.sexOfCalf}</td><td className="border border-emerald-100 px-2 py-2">{row.calfTag}</td><td className="border border-emerald-100 px-2 py-2">{row.dryDate}</td><td className="border border-emerald-100 px-2 py-2">{row.dryPeriod}</td><td className="border border-emerald-100 px-2 py-2">{row.tlmy}</td><td className="border border-emerald-100 px-2 py-2">{row.slmy}</td><td className="border border-emerald-100 px-2 py-2">{row.ll}</td><td className="border border-emerald-100 px-2 py-2">{row.py}</td><td className="border border-emerald-100 px-2 py-2">{row.sp}</td><td className="border border-emerald-100 px-2 py-2">{row.ci}</td></tr>)}</tbody></table></div>
                  </div>
                )}

                <div className="mt-4 flex justify-end"><button className="rounded-xl border border-emerald-300 bg-white px-4 py-2 text-emerald-800 hover:bg-emerald-50" onClick={goToNextFemaleTab}>Next Tab</button></div>
              </Section>
            )}

            {uiSelectedAnimal?.category === "Male" && uiSelectedAnimal?.isBreedingBull === "Yes" && (
              <Section title="Breeding Bull Tabs">
                <div className="mb-4 flex flex-wrap gap-3">
                  {MALE_TABS.map((tabName) => (
                    <button key={tabName} className={`rounded-xl border px-3 py-2 capitalize ${detailTab === tabName ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setDetailTab(tabName)}>
                      {tabName}
                    </button>
                  ))}
                </div>

                {detailTab === "photographs" && (
                  <AnimalPhotographsPanel photos={activeAnimal.maleDetails?.photographs || []} onUpload={handleAnimalPhotoUpload} onRemove={removeAnimalPhoto} />
                )}

                {detailTab === "pedigree" && (
                  <Grid>
                    <TextField label="Sire" value={activePedigree.sire} onChange={(v) => updateMalePedigree("sire", v)} />
                    <TextField label="Dam" value={activePedigree.dam} onChange={(v) => updateMalePedigree("dam", v)} />
                    <TextField label="Sire's sire" value={activePedigree.sireSire} onChange={(v) => updateMalePedigree("sireSire", v)} />
                    <TextField label="Sire's dam" value={activePedigree.sireDam} onChange={(v) => updateMalePedigree("sireDam", v)} />
                    <TextField label="Dam's sire" value={activePedigree.damSire} onChange={(v) => updateMalePedigree("damSire", v)} />
                    <TextField label="Dam's dam" value={activePedigree.damDam} onChange={(v) => updateMalePedigree("damDam", v)} />
                    <TextField label="Great-grandsire (SSS)" value={activePedigree.sireSireSire} onChange={(v) => updateMalePedigree("sireSireSire", v)} />
                    <TextField label="Great-granddam (SSD)" value={activePedigree.sireSireDam} onChange={(v) => updateMalePedigree("sireSireDam", v)} />
                    <TextField label="Great-grandsire (SDS)" value={activePedigree.sireDamSire} onChange={(v) => updateMalePedigree("sireDamSire", v)} />
                    <TextField label="Great-granddam (SDD)" value={activePedigree.sireDamDam} onChange={(v) => updateMalePedigree("sireDamDam", v)} />
                    <TextField label="Great-grandsire (DSS)" value={activePedigree.damSireSire} onChange={(v) => updateMalePedigree("damSireSire", v)} />
                    <TextField label="Great-granddam (DSD)" value={activePedigree.damSireDam} onChange={(v) => updateMalePedigree("damSireDam", v)} />
                    <TextField label="Great-grandsire (DDS)" value={activePedigree.damDamSire} onChange={(v) => updateMalePedigree("damDamSire", v)} />
                    <TextField label="Great-granddam (DDD)" value={activePedigree.damDamDam} onChange={(v) => updateMalePedigree("damDamDam", v)} />
                  </Grid>
                )}

                {detailTab === "disease testing" && (
                  <div className="space-y-5">
                    {(activeAnimal.maleDetails.diseaseTests || []).map((rec, idx) => (
                      <div key={`dt-${idx}`} className="rounded-2xl border border-emerald-100 bg-white p-3">
                        <Grid>
                          <DateField label="Testing date" value={rec.testDate || ""} onChange={(v) => updateDiseaseTest(idx, "testDate", v)} />
                          <TextField label="Disease / test" value={rec.testName || ""} onChange={(v) => updateDiseaseTest(idx, "testName", v)} />
                          <TextField label="Result" value={rec.result || ""} onChange={(v) => updateDiseaseTest(idx, "result", v)} />
                          <AreaField label="Remarks" value={rec.remarks || ""} onChange={(v) => updateDiseaseTest(idx, "remarks", v)} rows={2} />
                        </Grid>
                      </div>
                    ))}
                    <div className="flex gap-3"><button className="flex h-14 w-14 items-center justify-center rounded-2xl bg-emerald-600 text-4xl font-bold text-white hover:bg-emerald-700" onClick={addDiseaseTest}>+</button><button className="flex h-14 w-14 items-center justify-center rounded-2xl border border-slate-300 bg-white text-4xl font-bold text-slate-700 hover:bg-slate-50" onClick={removeDiseaseTest}>−</button></div>
                  </div>
                )}

                {detailTab === "dam record" && (
                  <div className="space-y-5">
                    <div className="rounded-2xl border border-emerald-100 bg-white p-4 shadow-sm">
                      <div className="mb-3 text-base font-semibold text-emerald-900">Dam record</div>
                      {!maleDamRecord ? (
                        <div className="rounded-2xl border border-dashed border-emerald-200 bg-emerald-50 px-4 py-6 text-sm text-slate-600">
                          Enter the dam tag in the male pedigree tab. When a matching female record exists in current or archived animals, her lactation-wise record will appear here automatically.
                        </div>
                      ) : (
                        <div className="space-y-3">
                          <div className="text-sm text-slate-600">Dam tag: <span className="font-semibold text-emerald-900">{maleDamRecord.damTag}</span></div>
                          <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70">
                            <table className="min-w-full border-collapse text-sm">
                              <thead className="bg-slate-50/90">
                                <tr className="bg-emerald-50 text-left text-emerald-900">
                                  <th className="border border-emerald-100 px-2 py-2 align-top" rowSpan={2}>Parameter(s)</th>
                                  <th className="border border-emerald-100 px-2 py-2 text-center" colSpan={10}>Lactation number</th>
                                </tr>
                                <tr className="bg-emerald-50 text-left text-emerald-900">
                                  {Array.from({ length: 10 }, (_, idx) => <th key={`dam-head-${idx + 1}`} className="border border-emerald-100 px-2 py-2 text-center">{idx + 1}</th>)}
                                </tr>
                              </thead>
                              <tbody>
                                {[
                                  ["Total Lactation Milk Yield", maleDamRecord.tlmy],
                                  ["Standard Lactation Milk Yield", maleDamRecord.slmy],
                                  ["Peak Yield", maleDamRecord.py],
                                ].map(([label, values]) => (
                                  <tr key={label}>
                                    <td className="border border-emerald-100 px-2 py-2 font-medium text-slate-800">{label}</td>
                                    {values.map((value, idx) => <td key={`${label}-${idx}`} className="border border-emerald-100 px-2 py-2 text-center">{value || ""}</td>)}
                                  </tr>
                                ))}
                                <tr>
                                  <td className="border border-emerald-100 px-2 py-2 font-medium text-slate-800">Age at first calving (days)</td>
                                  <td className="border border-emerald-100 px-2 py-2 text-center">{maleDamRecord.afc || ""}</td>
                                  <td className="border border-emerald-100 px-2 py-2 bg-emerald-50/60" colSpan={9}></td>
                                </tr>
                                <tr>
                                  <td className="border border-emerald-100 px-2 py-2 font-medium text-slate-800">Service periods (days)</td>
                                  <td className="border border-emerald-100 px-2 py-2 bg-emerald-50/60"></td>
                                  {maleDamRecord.sp.slice(1).map((value, idx) => <td key={`sp-${idx}`} className="border border-emerald-100 px-2 py-2 text-center">{value || ""}</td>)}
                                </tr>
                                <tr>
                                  <td className="border border-emerald-100 px-2 py-2 font-medium text-slate-800">Calving intervals of Dam (days)</td>
                                  <td className="border border-emerald-100 px-2 py-2 bg-emerald-50/60"></td>
                                  {maleDamRecord.ci.slice(1).map((value, idx) => <td key={`ci-${idx}`} className="border border-emerald-100 px-2 py-2 text-center">{value || ""}</td>)}
                                </tr>
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {detailTab === "ai information" && (
                  <div className="space-y-5">
                    <div className="flex flex-wrap items-end gap-4">
                      <div className="min-w-[220px]">
                        <label className="mb-1 block text-sm font-medium text-slate-700">AI Year Filter</label>
                        <select
                          className="w-full rounded-xl border border-emerald-200 bg-white px-3 py-2"
                          value={selectedBullAIYear}
                          onChange={(e) => setSelectedBullAIYear(e.target.value)}
                        >
                          {(bullAIYearOptions || ["All"]).map((year) => (
                            <option key={year} value={year}>
                              {year === "All" ? "All Years" : year}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>
                    <div className="overflow-x-auto rounded-2xl border border-emerald-100 bg-white p-3">
                      <table className="min-w-full border-collapse text-sm">
                        <thead>
                          <tr className="bg-emerald-50 text-left text-emerald-900">
                            {[
                              "Total AI",
                              "Conceived",
                              "Negative",
                              "Pending",
                              "CR %",
                              "Conceived in 1st attempt",
                              "FSCR %",
                            ].map((h) => (
                              <th key={h} className="border border-emerald-100 px-3 py-2 font-semibold">{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.totalAI}</td>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.conceived}</td>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.negative}</td>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.pending}</td>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.crPct}</td>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.firstAttemptConceived}</td>
                            <td className="border border-emerald-100 px-3 py-2">{selectedBullAIStats.firstServiceConceptionRate}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3">
                      <table className="min-w-full border-collapse text-sm">
                        <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Female Tag","Parity","AI No.","AI Date","Set No.","Result","Conception Date"].map((h) => <th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                        <tbody>
                          {selectedBullAIEvents.length === 0 ? (
                            <tr><td colSpan={7} className="border border-emerald-100 px-2 py-3 text-center text-slate-500">No inseminations linked to this bull for the selected filter.</td></tr>
                          ) : selectedBullAIEvents.map((row, idx) => (
                            <tr key={`bull-ai-${idx}`}>
                              <td className="border border-emerald-100 px-2 py-2">{row.femaleTag}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.parityNo}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.bullAiNo}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.aiDate}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.setNo || "—"}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.result}</td>
                              <td className="border border-emerald-100 px-2 py-2">{row.conceptionDate || "—"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {detailTab === "progenies born" && (
                  <div className="space-y-5">
                    <div className="flex gap-2">
                      <button className={`rounded-xl border px-3 py-2 ${maleProgenySubTab === "female" ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setMaleProgenySubTab("female")}>Female progenies</button>
                      <button className={`rounded-xl border px-3 py-2 ${maleProgenySubTab === "male" ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setMaleProgenySubTab("male")}>Male progenies</button>
                    </div>
                    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3">
                      <table className="min-w-full border-collapse text-sm">
                        <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Tag No.", "DOB", "Breed", "Status", "Archived?"].map((h) => <th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                        <tbody>
                          {(maleProgenySubTab === "female" ? femaleProgenies : maleProgenies).map((p) => (
                            <tr key={p.id}><td className="border border-emerald-100 px-2 py-2">{p.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{p.dob}</td><td className="border border-emerald-100 px-2 py-2">{p.breed}</td><td className="border border-emerald-100 px-2 py-2">{p.category === "Female" ? getFemaleLifecycle(p) : "Male"}</td><td className="border border-emerald-100 px-2 py-2">{isArchivedAnimal(p) ? "Yes" : "No"}</td></tr>
                          ))}
                          {(maleProgenySubTab === "female" ? femaleProgenies : maleProgenies).length === 0 && <tr><td colSpan={5} className="border border-emerald-100 px-2 py-3 text-center text-slate-500">No progenies linked to this bull yet.</td></tr>}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {detailTab === "performance of daughters" && (
                  <div className="space-y-5">
                    <div className="flex gap-2">
                      <button className={`rounded-xl border px-3 py-2 ${daughterPerfSubTab === "production" ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setDaughterPerfSubTab("production")}>Production</button>
                      <button className={`rounded-xl border px-3 py-2 ${daughterPerfSubTab === "reproduction" ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setDaughterPerfSubTab("reproduction")}>Reproduction</button>
                    </div>
                    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3">
                      {daughterPerfSubTab === "production" ? (
                        <table className="min-w-full border-collapse text-sm">
                          <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Daughter", "Current category", "Parity count", "Latest TLMY", "Latest SLMY", "Peak yield"].map((h) => <th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                          <tbody>
                            {femaleProgenies.map((d) => {
                              const lactations = d.femaleDetails.productionLactations || [];
                              const latest = lactations.length ? lactations[lactations.length - 1] : null;
                              const metrics = latest ? computeProductionMetrics(latest) : { totalLactationMilk: 0, standardLactationMilk: 0, peakYield: 0 };
                              return <tr key={d.id}><td className="border border-emerald-100 px-2 py-2">{d.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{getFemaleLifecycle(d)}</td><td className="border border-emerald-100 px-2 py-2">{d.femaleDetails.calvingParities.filter((p) => p.calvingDate).length}</td><td className="border border-emerald-100 px-2 py-2">{latest ? metrics.totalLactationMilk.toFixed(1) : ""}</td><td className="border border-emerald-100 px-2 py-2">{latest ? metrics.standardLactationMilk.toFixed(1) : ""}</td><td className="border border-emerald-100 px-2 py-2">{latest ? metrics.peakYield.toFixed(1) : ""}</td></tr>;
                            })}
                            {femaleProgenies.length === 0 && <tr><td colSpan={8} className="border border-emerald-100 px-2 py-3 text-center text-slate-500">No female progenies available yet.</td></tr>}
                          </tbody>
                        </table>
                      ) : (
                        <table className="min-w-full border-collapse text-sm">
                          <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Daughter", "AFC (days)", "Service periods (days)", "Calving Intervals (days)"].map((h) => <th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                          <tbody>
                            {femaleProgenies.map((d) => {
                              const calvingParities = (d.femaleDetails?.calvingParities || []).filter((p) => p.calvingDate).sort((a, b) => Number(a.parityNo) - Number(b.parityNo));
                              const afc = computeCalvingMetrics(d, 1).afc ?? "";
                              const servicePeriods = calvingParities
                                .map((cp) => computeCalvingMetrics(d, cp.parityNo).servicePeriod)
                                .filter((v) => v !== null && v !== undefined && v !== "");
                              const calvingIntervals = calvingParities
                                .map((cp) => computeCalvingMetrics(d, cp.parityNo).calvingInterval)
                                .filter((v) => v !== null && v !== undefined && v !== "");
                              return <tr key={d.id}><td className="border border-emerald-100 px-2 py-2">{d.tagNo}</td><td className="border border-emerald-100 px-2 py-2">{afc}</td><td className="border border-emerald-100 px-2 py-2">{servicePeriods.join(", ")}</td><td className="border border-emerald-100 px-2 py-2">{calvingIntervals.join(", ")}</td></tr>;
                            })}
                            {femaleProgenies.length === 0 && <tr><td colSpan={4} className="border border-emerald-100 px-2 py-3 text-center text-slate-500">No female progenies available yet.</td></tr>}
                          </tbody>
                        </table>
                      )}
                    </div>
                  </div>
                )}

                {detailTab === "health" && selectedHealth && (
                  <div className="space-y-5">
                    <div className="flex flex-wrap gap-3">{HEALTH_SUBTABS.map((sub) => <button key={sub.id} className={`rounded-xl border px-3 py-2 ${healthSubTab === sub.id ? "border-emerald-600 bg-emerald-600 text-white" : "border-emerald-200 bg-white text-emerald-800 hover:bg-emerald-50"}`} onClick={() => setHealthSubTab(sub.id)}>{sub.label}</button>)}</div>
                    {healthSubTab === "bodyWeight" && <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Body weight and recording date</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(maleTarget, "bodyWeightRecords", makeBodyWeightRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(maleTarget, "bodyWeightRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.bodyWeightRecords.map((rec, idx) => <div key={`mbw-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Recording date" value={rec.recordDate || ""} onChange={(v) => updateHealthRecord(maleTarget, "bodyWeightRecords", idx, "recordDate", v)} /><TextField label="Body weight" value={rec.bodyWeight || ""} onChange={(v) => updateHealthRecord(maleTarget, "bodyWeightRecords", idx, "bodyWeight", v)} /></Grid></div>)}</div></div>}
                    {healthSubTab === "deworming" && <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Deworming date and anthelmintic used</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(maleTarget, "dewormingRecords", makeDewormingRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(maleTarget, "dewormingRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.dewormingRecords.map((rec, idx) => <div key={`mdw-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Deworming date" value={rec.dewormingDate || ""} onChange={(v) => updateHealthRecord(maleTarget, "dewormingRecords", idx, "dewormingDate", v)} /><TextField label="Anthelmintic used" value={rec.anthelminticUsed || ""} onChange={(v) => updateHealthRecord(maleTarget, "dewormingRecords", idx, "anthelminticUsed", v)} /></Grid></div>)}</div></div>}
                    {healthSubTab === "vaccination" && <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Vaccination date and vaccine used</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(maleTarget, "vaccinationRecords", makeVaccinationRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(maleTarget, "vaccinationRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.vaccinationRecords.map((rec, idx) => <div key={`mv-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Vaccination date" value={rec.vaccinationDate || ""} onChange={(v) => updateHealthRecord(maleTarget, "vaccinationRecords", idx, "vaccinationDate", v)} /><TextField label="Vaccine used" value={rec.vaccineUsed || ""} onChange={(v) => updateHealthRecord(maleTarget, "vaccinationRecords", idx, "vaccineUsed", v)} /></Grid></div>)}</div></div>}
                    {healthSubTab === "treatment" && <div className="rounded-2xl border border-emerald-100 bg-white p-4"><div className="mb-2 flex items-center justify-between"><div className="text-base font-semibold text-emerald-900">Treatment dates, diagnosis and treatment given</div><div className="flex gap-3"><button className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-600 text-3xl font-bold text-white hover:bg-emerald-700" onClick={() => addHealthRecord(maleTarget, "treatmentRecords", makeTreatmentRecord())}>+</button><button className="flex h-12 w-12 items-center justify-center rounded-2xl border border-slate-300 bg-white text-3xl font-bold text-slate-700 hover:bg-slate-50" onClick={() => removeHealthRecord(maleTarget, "treatmentRecords")}>−</button></div></div><div className="grid gap-4">{selectedHealth.treatmentRecords.map((rec, idx) => <div key={`mt-${idx}`} className="rounded-2xl border border-emerald-100 bg-emerald-50 p-3"><Grid><DateField label="Treatment date" value={rec.treatmentDate || ""} onChange={(v) => updateHealthRecord(maleTarget, "treatmentRecords", idx, "treatmentDate", v)} /><TextField label="Diagnosis" value={rec.diagnosis || ""} onChange={(v) => updateHealthRecord(maleTarget, "treatmentRecords", idx, "diagnosis", v)} /><TextField label="Treatment given" value={rec.treatmentGiven || ""} onChange={(v) => updateHealthRecord(maleTarget, "treatmentRecords", idx, "treatmentGiven", v)} /></Grid></div>)}</div></div>}
                  </div>
                )}

                {detailTab === "overall history sheet" && (
                  <div className="space-y-5">
                    <div className="flex justify-end">
                      <button
                        className="rounded-xl border border-emerald-300 bg-white px-4 py-2 text-sm font-medium text-emerald-800 hover:bg-emerald-50"
                        onClick={() => exportBreedingBullHistoryPdf(activeAnimal, { femaleProgenies, maleProgenies, selectedBullAIEvents, selectedBullAIStats })}
                      >
                        Export History PDF
                      </button>
                    </div>
                    <Grid>
                      <TextField label="Breeding set" value={activeAnimal.breedingSet || ""} onChange={() => {}} readOnly />
                      <TextField label="Current set" value={activeAnimal.currentSet || "Yes"} onChange={() => {}} readOnly />
                      <TextField label="PT bull" value={activeAnimal.currentSet === "No" ? (activeAnimal.isPtBull || "No") : "—"} onChange={() => {}} readOnly />
                      <TextField label="Female progenies" value={String(femaleProgenies.length)} onChange={() => {}} readOnly />
                      <TextField label="Male progenies" value={String(maleProgenies.length)} onChange={() => {}} readOnly />
                    </Grid>
                    <AreaField label="Remarks" value={activeAnimal.maleDetails.historyMeta.remarks || ""} onChange={(v) => updateMaleHistoryMeta("remarks", v)} />
                    <TextField label="Book value" value={activeAnimal.maleDetails.historyMeta.bookValue || ""} onChange={(v) => updateMaleHistoryMeta("bookValue", v)} />
                    <div className="overflow-x-auto rounded-[24px] border border-slate-100 bg-white/70 rounded-2xl border border-emerald-100 bg-white p-3">
                      <table className="min-w-full border-collapse text-sm">
                        <thead className="bg-slate-50/90"><tr className="bg-emerald-50 text-left text-emerald-900">{["Disease test date", "Disease/test", "Result", "Remarks"].map((h) => <th key={h} className="border border-emerald-100 px-2 py-2">{h}</th>)}</tr></thead>
                        <tbody>
                          {activeAnimal.maleDetails.diseaseTests.map((r, idx) => <tr key={`ht-${idx}`}><td className="border border-emerald-100 px-2 py-2">{r.testDate}</td><td className="border border-emerald-100 px-2 py-2">{r.testName}</td><td className="border border-emerald-100 px-2 py-2">{r.result}</td><td className="border border-emerald-100 px-2 py-2">{r.remarks}</td></tr>)}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                <div className="mt-4 flex justify-end"><button className="rounded-xl border border-emerald-300 bg-white px-4 py-2 text-emerald-800 hover:bg-emerald-50" onClick={() => setDetailTab(nextDetailTab(detailTab, MALE_TABS))}>Next Tab</button></div>
              </Section>
            )}
          </div>
        </div>
        )}
          </div>
        </div>
      </div>
    </div>
  );
}

