import type { FUniver } from "@univerjs/core/facade";

const MARK_DIRTY_ROW_AUTO_HEIGHT_ID = "sheet.operation.mark-dirty-row-auto-height";
const SET_RANGE_VALUES_MUTATION_ID = "sheet.mutation.set-range-values";

interface CellValueMatrix {
  [row: string]:
    | {
        [col: string]: unknown;
      }
    | undefined;
}

interface SetRangeValuesMutationParams {
  unitId?: string;
  subUnitId?: string;
  cellValue?: CellValueMatrix;
}

function collectRowsFromCellValue(cellValue: CellValueMatrix | undefined): number[] {
  if (!cellValue) {
    return [];
  }

  const rows: number[] = [];
  for (const rowKey of Object.keys(cellValue)) {
    const row = Number(rowKey);
    if (Number.isInteger(row) && row >= 0 && cellValue[rowKey]) {
      rows.push(row);
    }
  }
  return rows;
}

function markRowsAutoHeight(
  univerAPI: FUniver,
  unitId: string,
  subUnitId: string,
  rows: number[],
): void {
  if (rows.length === 0) {
    return;
  }

  const ranges = rows.map((row) => ({
    startRow: row,
    endRow: row,
    startColumn: 0,
    endColumn: 0,
  }));

  // Undo 스택에 넣지 않고 idle 때 행 전체 기준으로 자동 높이를 다시 잰다.
  // 수식 결과는 mutation만 적용되어 행 높이가 비고,
  // 이후 다른 셀 편집 시 좁은 범위 재계산으로 높이가 갑자기 줄어들 수 있다.
  univerAPI.syncExecuteCommand(
    MARK_DIRTY_ROW_AUTO_HEIGHT_ID,
    {
      unitId,
      subUnitId,
      ranges,
      id: crypto.randomUUID(),
    },
    { onlyLocal: true },
  );
}

/**
 * 셀 값/수식 결과 반영 후 영향 받은 행의 자동 높이를 다시 맞춘다.
 * Univer 기본 동작은 수식 결과 때 행 높이를 갱신하지 않고,
 * 이후 편집 확정 때 편집 셀만 보고 줄일 수 있다.
 */
export function installRowAutoHeightFix(univerAPI: FUniver): () => void {
  const disposable = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
    if (event.id !== SET_RANGE_VALUES_MUTATION_ID) {
      return;
    }

    const params = event.params as SetRangeValuesMutationParams | undefined;
    if (!params?.unitId || !params.subUnitId) {
      return;
    }

    const rows = collectRowsFromCellValue(params.cellValue);
    if (rows.length === 0) {
      return;
    }

    markRowsAutoHeight(univerAPI, params.unitId, params.subUnitId, rows);
  });

  // 수식 비동기 계산 결과가 셀에 실제로 적용된 직후에 영향을 받은 행들의 높이를 다시 맞춥니다.
  const formulaDisposable = univerAPI.getFormula().calculationResultApplied((result) => {
    if (!result || !result.unitData) {
      return;
    }

    for (const unitId of Object.keys(result.unitData)) {
      const unitDataForSheet = result.unitData[unitId];
      if (!unitDataForSheet) {
        continue;
      }

      for (const subUnitId of Object.keys(unitDataForSheet)) {
        const matrix = unitDataForSheet[subUnitId];
        if (!matrix) {
          continue;
        }

        const rows: number[] = [];
        for (const rowKey of Object.keys(matrix)) {
          const row = Number(rowKey);
          if (Number.isInteger(row) && row >= 0) {
            rows.push(row);
          }
        }

        if (rows.length > 0) {
          markRowsAutoHeight(univerAPI, unitId, subUnitId, rows);
        }
      }
    }
  });

  return () => {
    disposable.dispose();
    formulaDisposable.dispose();
  };
}
