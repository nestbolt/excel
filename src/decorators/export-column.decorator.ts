import "reflect-metadata";
import { EXPORT_COLUMNS_META } from "./constants";
import type { ExportColumnOptions } from "./interfaces";

export function ExportColumn(opts?: ExportColumnOptions): PropertyDecorator {
  return (target: Object, propertyKey: string | symbol) => {
    const key = String(propertyKey);
    const ctor = target.constructor;

    const existing: Map<string, ExportColumnOptions> =
      Reflect.getOwnMetadata(EXPORT_COLUMNS_META, ctor) ?? new Map();

    existing.set(key, opts ?? {});
    Reflect.defineMetadata(EXPORT_COLUMNS_META, existing, ctor);
  };
}
