import { examples, type ExampleKey } from "$lib/examples";
import { highlight } from "$lib/server/highlight";
import type { PageServerLoad } from "./$types";

const KEYS: ExampleKey[] = [
  "recipeMailMerge",
  "recipeStyledBase",
  "recipeTrackedChanges",
  "previewEmbed",
];

export const load: PageServerLoad = async () => {
  const recipes = await Promise.all(
    KEYS.map(async (key) => {
      const ex = examples[key];
      return {
        key,
        title: ex.title,
        description: ex.description,
        path: ex.path,
        html: await highlight(ex.source, "ts"),
      };
    }),
  );
  return { recipes };
};
