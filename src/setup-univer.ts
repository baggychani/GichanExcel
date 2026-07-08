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
import { WrapStrategy } from "@univerjs/core";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import type { FUniver } from "@univerjs/core/facade";

import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";
import "@univerjs/preset-sheets-find-replace/lib/index.css";
import "@univerjs/preset-sheets-hyper-link/lib/index.css";
import "@univerjs/preset-sheets-note/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";

import { SsalmukTextSplitPlugin } from "./plugins/text-split";

export interface UniverApp {
  univerAPI: FUniver;
  dispose: () => void;
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
      }),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsNotePreset(),
    ],
    plugins: [SsalmukTextSplitPlugin],
  });

  univerAPI.createWorkbook({ name: "기찬엑셀" });
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet();
  // Google 스프레드시트처럼 A~Z(26열)만 사용
  sheet?.setColumnCount(26);
  // Ctrl+Enter 줄바꿈이 셀에 보이도록 기본 텍스트 줄바꿈 활성화
  sheet?.setDefaultStyle({ tb: WrapStrategy.WRAP });
  univerAPI.getFormula().executeCalculation();

  return {
    univerAPI,
    dispose: () => univer.dispose(),
  };
}
