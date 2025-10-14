import { CalendarSyncer } from "./CalendarSyncer";

declare global {
  // noinspection ES6ConvertVarToLetConst
  var main: () => void;
}

export function main(): void {
  new CalendarSyncer().syncCalendars();
}

globalThis.main = main;
