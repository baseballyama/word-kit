/**
 * The command registry — the single source of truth for every editor command.
 *
 * The coverage ledger validates that each WordprocessingML element classified
 * `edit` names a command id that exists here, so this registry is what makes
 * the completeness guarantee enforceable: an element can only claim to be
 * editable if a real command backs it.
 */

import { docpropsCommands } from "./docprops.js";
import { headerFooterCommands } from "./header-footer.js";
import { imageCommands } from "./image.js";
import { listCommands } from "./list.js";
import { numberingPropertyCommands } from "./numbering-props.js";
import { paragraphCommands } from "./paragraph.js";
import { paragraphPropertyCommands, runPropertyCommands } from "./properties.js";
import { rawCommands } from "./raw.js";
import { referencesCommands } from "./references.js";
import { reviewCommands } from "./review.js";
import { sectionCommands } from "./section.js";
import { sectionPropertyCommands } from "./section-props.js";
import { settingsCommands } from "./settings.js";
import { structureCommands } from "./structure.js";
import { styleCommands } from "./style.js";
import { stylePropertyCommands } from "./style-props.js";
import { tableCommands } from "./table.js";
import { tablePropertyCommands } from "./table-props.js";
import { textCommands } from "./text.js";
import type { Command } from "./types.js";

/** Every command, in ribbon order. */
export const ALL_COMMANDS: ReadonlyArray<Command<never>> = [
  ...textCommands,
  ...runPropertyCommands,
  ...paragraphCommands,
  ...paragraphPropertyCommands,
  ...structureCommands,
  ...listCommands,
  ...numberingPropertyCommands,
  ...tableCommands,
  ...tablePropertyCommands,
  ...imageCommands,
  ...styleCommands,
  ...stylePropertyCommands,
  ...sectionCommands,
  ...sectionPropertyCommands,
  ...headerFooterCommands,
  ...referencesCommands,
  ...reviewCommands,
  ...docpropsCommands,
  ...settingsCommands,
  ...rawCommands,
] as ReadonlyArray<Command<never>>;

/** The set of all command ids, for coverage validation. */
export const COMMAND_IDS: ReadonlySet<string> = new Set(ALL_COMMANDS.map((c) => c.id));

/** Look up a command by id. */
export function getCommand(id: string): Command<never> | undefined {
  return ALL_COMMANDS.find((c) => c.id === id);
}

/** All commands in a feature group, for building ribbon sections. */
export function commandsInGroup(group: Command<never>["group"]): Command<never>[] {
  return ALL_COMMANDS.filter((c) => c.group === group);
}
