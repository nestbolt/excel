import { describe, it, expect } from "vitest";
import { validateRow, mapDtoErrors } from "../src/helpers/validate-row";
import type { ValidationRules } from "../src/concerns/with-validation.interface";

describe("validateRow", () => {
  it("should skip undefined rule entries", async () => {
    const rules: ValidationRules = {
      name: [{ validate: (v) => v === "ok", message: "fail" }],
      age: undefined, // should be skipped gracefully
    };

    const result = await validateRow({ name: "ok", age: 10 }, rules, 1);
    expect(result).toBeNull();
  });
});

describe("mapDtoErrors", () => {
  it("should return null for empty errors array", () => {
    expect(mapDtoErrors([], 1)).toBeNull();
  });

  it("should map errors with constraints", () => {
    const errors = [
      { property: "email", constraints: { isEmail: "email must be valid" } },
    ];
    const result = mapDtoErrors(errors, 3);
    expect(result).not.toBeNull();
    expect(result!.row).toBe(3);
    expect(result!.errors[0].field).toBe("email");
    expect(result!.errors[0].messages).toEqual(["email must be valid"]);
  });

  it("should return empty messages when constraints is undefined", () => {
    const errors = [{ property: "name", constraints: undefined }];
    const result = mapDtoErrors(errors, 2);
    expect(result).not.toBeNull();
    expect(result!.errors[0].field).toBe("name");
    expect(result!.errors[0].messages).toEqual([]);
  });

  it("should return empty messages when constraints is null", () => {
    const errors = [{ property: "age", constraints: null }];
    const result = mapDtoErrors(errors, 5);
    expect(result).not.toBeNull();
    expect(result!.errors[0].field).toBe("age");
    expect(result!.errors[0].messages).toEqual([]);
  });
});
