import type { ValidationRules } from "../concerns/with-validation.interface";
import type {
  FieldError,
  ImportValidationError,
} from "../interfaces/import-result.interface";

export async function validateRow(
  row: Record<string, any>,
  rulesOrDto: ValidationRules | { dto: new (...args: any[]) => any },
  rowNumber: number,
): Promise<ImportValidationError | null> {
  if ("dto" in rulesOrDto && typeof rulesOrDto.dto === "function") {
    return validateWithDto(row, rulesOrDto.dto, rowNumber);
  }
  return validateWithRules(row, rulesOrDto as ValidationRules, rowNumber);
}

function validateWithRules(
  row: Record<string, any>,
  rules: ValidationRules,
  rowNumber: number,
): ImportValidationError | null {
  const fieldErrors: FieldError[] = [];

  for (const [field, fieldRules] of Object.entries(rules)) {
    if (!fieldRules) continue;
    const messages: string[] = [];
    for (const rule of fieldRules) {
      if (!rule.validate(row[field], row)) {
        messages.push(rule.message);
      }
    }
    if (messages.length > 0) {
      fieldErrors.push({ field, messages });
    }
  }

  return fieldErrors.length > 0
    ? { row: rowNumber, errors: fieldErrors }
    : null;
}

let cachedValidator: any;
let cachedTransformer: any;

async function loadDtoDeps(): Promise<{ validator: any; transformer: any }> {
  if (cachedValidator && cachedTransformer) {
    return { validator: cachedValidator, transformer: cachedTransformer };
  }
  try {
    cachedValidator = await import("class-validator");
    cachedTransformer = await import("class-transformer");
  } catch /* v8 ignore next 4 */ {
    throw new Error(
      "WithValidation with DTO requires class-validator and class-transformer. " +
        "Install them: pnpm add class-validator class-transformer",
    );
  }
  return { validator: cachedValidator, transformer: cachedTransformer };
}

async function validateWithDto(
  row: Record<string, any>,
  dto: new (...args: any[]) => any,
  rowNumber: number,
): Promise<ImportValidationError | null> {
  const { validator, transformer } = await loadDtoDeps();

  const instance = transformer.plainToInstance(dto, row);
  const errors: any[] = validator.validateSync(instance);

  return mapDtoErrors(errors, rowNumber);
}

/** @internal Exported for testing only. */
export function mapDtoErrors(
  errors: any[],
  rowNumber: number,
): ImportValidationError | null {
  if (errors.length === 0) return null;

  const fieldErrors: FieldError[] = errors.map((err: any) => ({
    field: err.property,
    messages: err.constraints
      ? (Object.values(err.constraints) as string[])
      : [],
  }));

  return { row: rowNumber, errors: fieldErrors };
}
