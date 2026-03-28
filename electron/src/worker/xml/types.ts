/** 공유 타입, 인터페이스, 상수 */

export type ComHandle = unknown;

export interface ComBridge {
  comGet(handle: ComHandle, prop: string): unknown;
  comPut(handle: ComHandle, prop: string, value: unknown): void;
  comCallWith(handle: ComHandle, method: string, args: unknown[]): unknown;
}

export interface CellInfo {
  col: number;
  row: number;
  colSpan: number;
  rowSpan: number;
  text: string;
}

export interface DocElement {
  type: 'p' | 'table';
  index: number;     // section 내 offset
  content: string;   // 원본 XML
}

export interface ElementLocation {
  xml: string;       // 해당 요소의 XML 문자열
  start: number;     // 전체 XML 내 절대 시작 위치
  end: number;       // 전체 XML 내 절대 끝 위치
}

export interface PageBoundary {
  page: number;
  startList: number;
  startPara: number;
  endList: number;
  endPara: number;
}

/** structureDetail용 화이트리스트 — 태그별 유지할 속성 */
export const ATTR_WHITELIST: Record<string, string[] | '*'> = {
  TABLE:      ['ColCount', 'RowCount', 'BorderFill'],
  ROW:        [],
  CELL:       ['ColAddr', 'RowAddr', 'ColSpan', 'RowSpan', 'Width', 'Height', 'BorderFill'],
  P:          ['ParaShape', 'Style'],
  TEXT:       ['CharShape'],
  CHAR:       [],
  LINEBREAK:  [],
  PICTURE:    '*',
  SHAPEOBJECT: ['InstId'],
  SIZE:       ['Width', 'Height', 'WidthRelTo', 'HeightRelTo'],
  POSITION:   ['TreatAsChar', 'HorzAlign', 'VertAlign'],
  IMAGE:      '*',
  CAPTION:    ['Side'],
  PARALIST:   ['VertAlign'],
};
