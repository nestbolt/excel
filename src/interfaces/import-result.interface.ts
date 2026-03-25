export interface ImportResult<T = any> {
  rows: T[];
  errors: ImportValidationError[];
  skipped: number;
}

export interface ImportValidationError {
  row: number;
  errors: FieldError[];
}

export interface FieldError {
  field: string;
  messages: string[];
}
