import type { FUniver } from "@univerjs/core/facade";
import { SplitDelimiterEnum } from "@univerjs/sheets";

export type DelimiterMode = "tab" | "spaces" | "comma" | "custom";

export interface TextToColumnsOptions {
  mode: DelimiterMode;
  customDelimiter: string;
  treatMultipleDelimitersAsOne: boolean;
}

function resolveDelimiter(options: TextToColumnsOptions): {
  delimiter: SplitDelimiterEnum;
  customDelimiter?: string;
} {
  if (options.mode === "custom") {
    const customDelimiter = options.customDelimiter.trim();
    if (!customDelimiter) {
      throw new Error("직접 입력 구분자를 입력해 주세요.");
    }

    return {
      delimiter: SplitDelimiterEnum.Custom,
      customDelimiter,
    };
  }

  if (options.mode === "tab") {
    return { delimiter: SplitDelimiterEnum.Tab };
  }

  if (options.mode === "comma") {
    return { delimiter: SplitDelimiterEnum.Comma };
  }

  return { delimiter: SplitDelimiterEnum.Space };
}

export function splitActiveRange(
  univerAPI: FUniver,
  options: TextToColumnsOptions,
): string {
  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) {
    throw new Error("열린 통합 문서가 없습니다.");
  }

  const activeRange = workbook.getActiveRange();
  if (!activeRange) {
    throw new Error("먼저 분할할 셀이나 범위를 선택해 주세요.");
  }

  const width = activeRange.getWidth();
  if (width > 1) {
    throw new Error("한 열만 선택한 뒤 분할해 주세요.");
  }

  const { delimiter, customDelimiter } = resolveDelimiter(options);

  (
    activeRange as typeof activeRange & {
      splitTextToColumns: (
        treatMultipleDelimitersAsOne?: boolean,
        delimiter?: SplitDelimiterEnum,
        customDelimiter?: string,
      ) => void;
    }
  ).splitTextToColumns(
    options.treatMultipleDelimitersAsOne,
    delimiter,
    customDelimiter,
  );

  return activeRange.getA1Notation();
}
