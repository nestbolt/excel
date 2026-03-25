/**
 * Skip invalid rows instead of throwing during validation.
 *
 * Marker concern — set `skipsOnError` to `true`.
 */
export interface SkipsOnError {
  readonly skipsOnError: true;
}
