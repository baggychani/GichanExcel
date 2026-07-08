import { CommandType, type ICommand } from "@univerjs/core";
import {
  OPEN_TEXT_SPLIT_DIALOG_COMMAND_ID,
  TEXT_SPLIT_EVENT,
} from "./constants";

export const OpenTextSplitDialogCommand: ICommand = {
  id: OPEN_TEXT_SPLIT_DIALOG_COMMAND_ID,
  type: CommandType.COMMAND,
  handler: () => {
    window.dispatchEvent(new CustomEvent(TEXT_SPLIT_EVENT));
    return true;
  },
};
