export type Row = Record<string, string | number | boolean | null>;

export type DataTable = {
  workbook: string;
  file: string;
  sheet: string;
  columns: string[];
  rows: Row[];
};

export type QueryPlan = {
  operation: "count_unique" | "count_rows" | "list" | "sum" | "average" | "min" | "max" | "group_count" | "lookup" | "unsupported";
  subject: string;
  target_column: string | null;
  group_by: string | null;
  preferred_sheet: string | null;
  filters: Array<{
    column: string;
    operator: "equals" | "contains";
    value: string;
  }>;
  interpretation: string;
};

export type QueryResult = {
  answer: string;
  interpretation: string;
  sources: string[];
  table?: Row[];
  warning?: string;
};
