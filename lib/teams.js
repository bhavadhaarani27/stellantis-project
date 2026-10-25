"use strict";
/*
 * Standard EHEQ team catalogue, taken from the Session 1 and Session 2 decks.
 * Teams are matched to deck sections by a normalised key, so month-to-month
 * differences in capitalisation or punctuation do not matter.
 */
const CATALOGUE = [
  { key: "EEQ DASHBOARD", label: "EEQ Dashboard" },
  { key: "EEQ STRATEGIC ACTION", label: "EEQ Strategic Action" },
  { key: "INFIELD QUALITY SA", label: "Infield Quality - SA" },
  { key: "INFIELD QUALITY IAP", label: "Infield Quality - IAP" },
  { key: "INFIELD QUALITY CHN", label: "Infield Quality - CHN" },
  { key: "INFIELD QUALITY NA", label: "Infield Quality - NA" },
  { key: "INFIELD QUALITY EMEA", label: "Infield Quality - EMEA" },
  { key: "TOPHAT QUALITY NEA AEE", label: "Tophat Quality - NEA/AEE" },
  { key: "TOPHAT QUALITY ATL H", label: "Tophat Quality - ATL H" },
  { key: "TOPHAT QUALITY ATL MID", label: "Tophat Quality - ATL MID" },
  { key: "TOPHAT QUALITY BRAIN", label: "Tophat Quality - BRAIN" },
  { key: "TOPHAT QUALITY SEAR", label: "Tophat Quality - SEAR" },
  { key: "FUNCTIONAL QUALITY EA EVIO", label: "Functional Quality - EA/EVIO (EEAV)" },
  { key: "FUNCTIONAL QUALITY HWE", label: "Functional Quality - HWE" },
  { key: "FUNCTIONAL QUALITY EEHD", label: "Functional Quality - EEHD" },
  { key: "2026 INDIVIDUAL TARGET", label: "2026 Individual Target" }
];

// "INFIELD QUALITY_NA" and "Infield Quality - NA" both become "INFIELD QUALITY NA".
function teamKey(name) {
  return String(name || "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// Long words are capitalised, short acronyms (EEQ, NA, TWS) stay upper-case.
function prettyName(name) {
  const entry = CATALOGUE.find((team) => team.key === teamKey(name));
  if (entry) return entry.label;
  return String(name).replace(/_+/g, " ").replace(/\s*-\s*/g, " ").replace(/\s+/g, " ").trim()
    .split(" ").map((w) => (/^[A-Z]{5,}$/.test(w) ? w[0] + w.slice(1).toLowerCase() : w)).join(" ");
}

// Merges the catalogue with the sections found in a file (file counts, plus extra sections the catalogue lacks).
function buildTeamList(fileSections) {
  const found = new Map(fileSections.map((s) => [teamKey(s.name), s]));
  const list = CATALOGUE.map((team) => {
    const section = found.get(team.key);
    return { key: team.key, label: section ? section.name : team.label, slides: section ? section.count : 0 };
  });
  for (const [key, section] of found) {
    if (key === "EHEQ" || CATALOGUE.some((team) => team.key === key)) continue;
    list.push({ key, label: prettyName(section.name), slides: section.count });
  }
  return list;
}

module.exports = { CATALOGUE, teamKey, prettyName, buildTeamList };
