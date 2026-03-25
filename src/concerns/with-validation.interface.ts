/**
 * Validate imported rows using custom rules or a class-validator DTO.
 *
 * Return an object with a `dto` key for class-validator integration,
 * or a `ValidationRules` map for custom rule functions.
 */
export interface WithValidation<T = Record<string, any>> {
  rules(): ValidationRules<T> | { dto: new (...args: any[]) => any };
}

export type ValidationRules<T = Record<string, any>> = {
  [K in keyof T]?: ValidationRule[];
};

export interface ValidationRule {
  validate: (value: any, row: Record<string, any>) => boolean;
  message: string;
}
