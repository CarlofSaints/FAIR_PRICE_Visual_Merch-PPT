export interface QAPair {
  question: string;
  answer: string;
  imageUrl?: string; // URL in the column immediately to the RIGHT of this question
}

export interface SurveySection {
  name: string;
  qaPairs: QAPair[];
}

export interface SurveyRow {
  id: string | number;
  email: string;
  firstName: string;
  lastName: string;
  fullName: string;
  store: string;
  storeCode: string;
  province: string;
  date: string;       // formatted: "18 Feb 2026"
  dateRaw: Date | null;
  dayOfWeek: string;  // e.g. "Tuesday"
  sections: SurveySection[];
  imageEntries: ImageEntry[]; // all images in this row
}

export interface ImageEntry {
  sectionName: string;
  imageUrl: string;
  imageHeader: string;    // header of the image column
  precedingQuestion: string; // Q column header to the left
  precedingAnswer: string;   // Q column answer to the left
}

export interface ParsedData {
  rows: SurveyRow[];
  uniqueUsers: string[];   // unique emails
  uniqueStores: string[];  // unique store names
  uniqueDays: string[];    // unique day-of-week values
  dateRange: { from: string; to: string };
  totalRows: number;
}

export interface UserSummary {
  fullName: string;
  email: string;
  uniqueStores: number;
  totalSurveys: number;
  dayCounts: Record<string, number>; // day → count
}
