import { examples } from "$lib/examples";
import { highlight } from "$lib/server/highlight";
import type { PageServerLoad } from "./$types";

export const load: PageServerLoad = async () => {
  const previewEmbed = await highlight(examples.previewEmbed.source, "ts");
  return {
    previewEmbed: { ...examples.previewEmbed, html: previewEmbed },
  };
};
