import { examples } from "$lib/examples";
import { highlight } from "$lib/server/highlight";
import type { PageServerLoad } from "./$types";

export const load: PageServerLoad = async () => {
  const fromScratch = await highlight(examples.fromScratch.source, "ts");
  const templateFill = await highlight(examples.templateFill.source, "ts");
  return {
    fromScratch: { ...examples.fromScratch, html: fromScratch },
    templateFill: { ...examples.templateFill, html: templateFill },
  };
};
