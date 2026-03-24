/**
 * Set workbook-level document properties.
 */
export interface WithProperties {
  properties(): ExcelProperties;
}

export interface ExcelProperties {
  creator?: string;
  lastModifiedBy?: string;
  title?: string;
  subject?: string;
  description?: string;
  keywords?: string;
  category?: string;
  company?: string;
  manager?: string;
}
