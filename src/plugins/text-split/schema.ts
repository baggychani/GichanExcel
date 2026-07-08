import { RibbonDataGroup, type MenuSchemaType } from "@univerjs/ui";
import { TEXT_SPLIT_TOOLBAR_MENU_ID } from "./constants";
import { TextSplitToolbarMenuItemFactory } from "./menu";

/** 데이터 탭 → ribbon.data.others → 「텍스트를 숫자로 변환」(order 0) 바로 옆 */
export const textSplitMenuSchema: MenuSchemaType = {
  [RibbonDataGroup.OTHERS]: {
    [TEXT_SPLIT_TOOLBAR_MENU_ID]: {
      order: 1,
      menuItemFactory: TextSplitToolbarMenuItemFactory,
    },
  },
};

