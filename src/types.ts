export type PlaceholderMap = Record<string, string>;
export type ImageMap = Record<string, Buffer>;

export interface OfficifyOptions {
  // Optional hint for input extension: ".odt", ".ods", ".odp"
  inputExtensionHint?: string;
  // Optional explicit soffice path override
  sofficePath?: string;
}
