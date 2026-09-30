// Registers the gap-filling plugins with HyperFormula (PRD §7.4). Each function here was confirmed
// missing or non-Excel-compatible in HyperFormula 3.4.0 during Milestone 1; see docs/milestones/M1.md.

import { HyperFormula } from 'hyperformula';
import { AverageIfsPlugin } from './averageifs.ts';
import { IndexPlugin } from './indexVector.ts';
import { LookupPlugin } from './lookup.ts';
import { RegressionPlugin } from './regression.ts';
import { TextPlugin } from './text.ts';

/** Plugins for functions HyperFormula lacks. */
const ADDED = [AverageIfsPlugin, LookupPlugin, RegressionPlugin];
/** Plugins that replace a built-in whose results differ from Excel's. */
const OVERRIDES = [TextPlugin, IndexPlugin];

const PLUGINS = [...ADDED, ...OVERRIDES];

/** Function names provided by plugins (aliases included), e.g. for verification.json. */
export const PLUGIN_FUNCTIONS: readonly string[] = PLUGINS.flatMap((plugin) => [
  ...Object.keys(plugin.implementedFunctions),
  ...Object.keys((plugin as { aliases?: Record<string, string> }).aliases ?? {}),
]).sort();

let registered = false;

/**
 * Idempotent. Registration is global to HyperFormula and applies to instances created afterwards.
 * Formulas are parsed in HyperFormula's default language (enGB), which needs a name for each function.
 */
export function registerPlugins(): void {
  if (registered) return;
  for (const plugin of OVERRIDES) {
    for (const name of Object.keys(plugin.implementedFunctions)) {
      HyperFormula.unregisterFunction(name);
    }
  }
  const enGB = Object.fromEntries(PLUGIN_FUNCTIONS.map((name) => [name, name]));
  for (const plugin of PLUGINS) HyperFormula.registerFunctionPlugin(plugin, { enGB });
  registered = true;
}
