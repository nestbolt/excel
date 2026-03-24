/**
 * Returned by `ExcelService.download()`.
 *
 * Contains everything needed to build an HTTP file-download response.
 */
export interface ExcelDownloadResult {
  /** The file contents. */
  buffer: Buffer;
  /** Suggested filename including extension. */
  filename: string;
  /** MIME content-type for the response header. */
  contentType: string;
}
