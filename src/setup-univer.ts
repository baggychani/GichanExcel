import { UniverSheetsConditionalFormattingPreset } from "@univerjs/preset-sheets-conditional-formatting";
import sheetsConditionalFormattingKoKR from "@univerjs/preset-sheets-conditional-formatting/locales/ko-KR";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreKoKR from "@univerjs/preset-sheets-core/locales/ko-KR";
import { UniverSheetsDataValidationPreset } from "@univerjs/preset-sheets-data-validation";
import sheetsDataValidationKoKR from "@univerjs/preset-sheets-data-validation/locales/ko-KR";
import { UniverSheetsFilterPreset } from "@univerjs/preset-sheets-filter";
import sheetsFilterKoKR from "@univerjs/preset-sheets-filter/locales/ko-KR";
import { UniverSheetsFindReplacePreset } from "@univerjs/preset-sheets-find-replace";
import sheetsFindReplaceKoKR from "@univerjs/preset-sheets-find-replace/locales/ko-KR";
import { UniverSheetsHyperLinkPreset } from "@univerjs/preset-sheets-hyper-link";
import sheetsHyperLinkKoKR from "@univerjs/preset-sheets-hyper-link/locales/ko-KR";
import { UniverSheetsNotePreset } from "@univerjs/preset-sheets-note";
import sheetsNoteKoKR from "@univerjs/preset-sheets-note/locales/ko-KR";
import { UniverSheetsSortPreset } from "@univerjs/preset-sheets-sort";
import sheetsSortKoKR from "@univerjs/preset-sheets-sort/locales/ko-KR";
import { CalculationMode } from "@univerjs/sheets-formula";
import { LocaleType, mergeLocales, WrapStrategy } from "@univerjs/core";
import type { FUniver } from "@univerjs/core/facade";

import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";
import "@univerjs/preset-sheets-find-replace/lib/index.css";
import "@univerjs/preset-sheets-hyper-link/lib/index.css";
import "@univerjs/preset-sheets-note/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";

import { createUniver } from "./lib/create-univer";
import { GichanTextSplitPlugin } from "./plugins/text-split";

// Windows에는 기본으로 깔려 있어 "폰트가 없습니다" 경고 없이 바로 쓸 수 있는
// 한글 시스템 폰트입니다. 그 외 폰트는 폰트 이름 입력창에 직접 타이핑해서
// 쓸 수 있습니다(설치돼 있으면 바로 적용됩니다).
const KOREAN_SYSTEM_FONTS = [
  { value: "Malgun Gothic", label: "맑은 고딕", category: "sans-serif" as const },
  { value: "Gulim", label: "굴림", category: "sans-serif" as const },
  { value: "Dotum", label: "돋움", category: "sans-serif" as const },
  { value: "Batang", label: "바탕", category: "serif" as const },
  { value: "Gungsuh", label: "궁서", category: "serif" as const },
];

export interface UniverApp {
  univerAPI: FUniver;
  dispose: () => void;
}

/**
 * 새로 만들어지거나 불러온 시트에 공통으로 적용할 기본 동작.
 * - Google 스프레드시트처럼 기본은 A~Z(26열)만 사용 (단, 더 많은 열을 가진
 *   파일을 열었을 때는 데이터가 잘리지 않도록 열을 줄이지는 않습니다)
 * - 셀 텍스트가 길어지면 옆 셀로 흘러넘치지 않고 줄바꿈되도록 기본 서식 지정
 * - 이미 들어있는 데이터(가져오기/자동저장 복구 등)의 행 높이도 줄바꿈 내용에 맞게 재계산
 */
export function applySheetDefaults(univerAPI: FUniver): void {
  const workbook = univerAPI.getActiveWorkbook();
  const sheet = workbook?.getActiveSheet();
  if (!sheet) {
    return;
  }

  sheet.setColumnCount(Math.max(26, sheet.getMaxColumns()));
  sheet.setDefaultStyle({ tb: WrapStrategy.WRAP });

  const dataRange = sheet.getDataRange();
  if (dataRange) {
    const range = dataRange.getRange();
    sheet.setRangesAutoHeight([range]);

    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const height = sheet.getRowHeight(row);
      if (height > 0) {
        sheet.setRowHeightsForced(row, 1, height);
      }
    }
  }
}

export function setupUniver(container: string): UniverApp {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.KO_KR,
    locales: {
      [LocaleType.KO_KR]: mergeLocales(
        sheetsCoreKoKR,
        sheetsFilterKoKR,
        sheetsSortKoKR,
        sheetsFindReplaceKoKR,
        sheetsConditionalFormattingKoKR,
        sheetsDataValidationKoKR,
        sheetsHyperLinkKoKR,
        sheetsNoteKoKR,
      ),
    },
    presets: [
      UniverSheetsCorePreset({
        container,
        formulaBar: true,
        // Tauri WebView에서 Worker RPC가 불안정해 수식 결과가 비는 경우가 있어
        // 메인 스레드에서 직접 계산합니다.
        formula: {
          initialFormulaComputing: CalculationMode.FORCED,
        },
        // 폰트 목록: 자주 쓰는 한글 시스템 폰트를 기본 목록 위에 추가합니다.
        // 목록에 없는 폰트도 입력창에 직접 타이핑하면 바로 적용됩니다.
        customFontFamily: KOREAN_SYSTEM_FONTS,
      }),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsNotePreset(),
    ],
    plugins: [GichanTextSplitPlugin],
  });

  univerAPI.createWorkbook({ name: "기찬엑셀" });
  applySheetDefaults(univerAPI);
  univerAPI.getFormula().executeCalculation();

  return {
    univerAPI,
    dispose: () => univer.dispose(),
  };
}
