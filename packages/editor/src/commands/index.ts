/** Barrel of all command objects and group arrays. */

export * from "./text.js";
export * from "./properties.js";
export * from "./docprops.js";
export * from "./settings.js";
export * from "./raw.js";
export * from "./paragraph.js";
export * from "./structure.js";
export * from "./list.js";
export * from "./table.js";
export * from "./table-props.js";
export * from "./style-props.js";
export * from "./numbering-props.js";
export * from "./section-props.js";
export * from "./image.js";
export * from "./style.js";
export * from "./section.js";
export * from "./header-footer.js";
export * from "./references.js";
export * from "./review.js";
export { ALL_COMMANDS, COMMAND_IDS, commandsInGroup, getCommand } from "./registry.js";
export { type Command, type FeatureGroup, runCommand } from "./types.js";
