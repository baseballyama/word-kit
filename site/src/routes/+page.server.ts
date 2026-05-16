import { examples, type ExampleKey } from "$lib/examples";
import { highlight } from "$lib/server/highlight";
import type { PageServerLoad } from "./$types";

const HERO_KEYS: ExampleKey[] = ["fromScratch", "previewEmbed"];

export const load: PageServerLoad = async () => {
  const hero = await Promise.all(
    HERO_KEYS.map(async (key) => {
      const ex = examples[key];
      return {
        key,
        title: ex.title,
        path: ex.path,
        description: ex.description,
        source: ex.source,
        html: await highlight(ex.source, "ts"),
      };
    }),
  );
  return { hero };
};
