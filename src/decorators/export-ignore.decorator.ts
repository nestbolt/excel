import "reflect-metadata";
import { EXPORT_IGNORE_META } from "./constants";

export function ExportIgnore(): PropertyDecorator {
  return (target: Object, propertyKey: string | symbol) => {
    const key = String(propertyKey);
    const ctor = target.constructor;

    const existing: Set<string> =
      Reflect.getOwnMetadata(EXPORT_IGNORE_META, ctor) ?? new Set();

    existing.add(key);
    Reflect.defineMetadata(EXPORT_IGNORE_META, existing, ctor);
  };
}
