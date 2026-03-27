"use strict";
/** 공유 타입, 인터페이스, 상수 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ATTR_WHITELIST = void 0;
/** structureDetail용 화이트리스트 — 태그별 유지할 속성 */
exports.ATTR_WHITELIST = {
    TABLE: ['ColCount', 'RowCount', 'BorderFill'],
    ROW: [],
    CELL: ['ColAddr', 'RowAddr', 'ColSpan', 'RowSpan', 'Width', 'Height', 'BorderFill'],
    P: ['ParaShape', 'Style'],
    TEXT: ['CharShape'],
    CHAR: [],
    LINEBREAK: [],
    PICTURE: '*',
    SHAPEOBJECT: ['InstId'],
    SIZE: ['Width', 'Height', 'WidthRelTo', 'HeightRelTo'],
    POSITION: ['TreatAsChar', 'HorzAlign', 'VertAlign'],
    IMAGE: '*',
    CAPTION: ['Side'],
    PARALIST: ['VertAlign'],
};
