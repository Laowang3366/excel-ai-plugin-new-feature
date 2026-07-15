import { z } from "zod";

import { USER_DATA_ERASE_CONFIRMATION } from "./userDataEraseContract";

export const EraseUserDataInput = z
  .object({
    confirmation: z.literal(USER_DATA_ERASE_CONFIRMATION),
  })
  .strict();
export type EraseUserDataInput = z.infer<typeof EraseUserDataInput>;
