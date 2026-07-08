import type { IAccessor } from "@univerjs/core";
import { UniverInstanceType } from "@univerjs/core";
import {
  RangeProtectionPermissionEditPoint,
  WorkbookEditablePermission,
  WorksheetEditPermission,
  WorksheetSetCellValuePermission,
} from "@univerjs/sheets";
import {
  getCurrentRangeDisable$,
  getObservableWithExclusiveRange$,
} from "@univerjs/sheets-ui";
import type { IMenuButtonItem } from "@univerjs/ui";
import { getMenuHiddenObservable, MenuItemType } from "@univerjs/ui";
import {
  OPEN_TEXT_SPLIT_DIALOG_COMMAND_ID,
  TEXT_SPLIT_TOOLBAR_MENU_ID,
} from "./constants";

export function TextSplitToolbarMenuItemFactory(
  accessor: IAccessor,
): IMenuButtonItem<string> {
  return {
    id: TEXT_SPLIT_TOOLBAR_MENU_ID,
    commandId: OPEN_TEXT_SPLIT_DIALOG_COMMAND_ID,
    type: MenuItemType.BUTTON,
    title: "텍스트를 열로 분할",
    disabled$: getObservableWithExclusiveRange$(
      accessor,
      getCurrentRangeDisable$(accessor, {
        workbookTypes: [WorkbookEditablePermission],
        worksheetTypes: [WorksheetEditPermission, WorksheetSetCellValuePermission],
        rangeTypes: [RangeProtectionPermissionEditPoint],
      }),
    ),
    hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET),
  };
}
