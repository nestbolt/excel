import "reflect-metadata";
import { EXPORTABLE_META } from "./constants";
import type { ExportableOptions } from "./interfaces";

export function Exportable(opts?: ExportableOptions): ClassDecorator {
  return (target: Function) => {
    Reflect.defineMetadata(EXPORTABLE_META, opts ?? {}, target);
  };
}
