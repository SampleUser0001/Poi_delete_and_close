package ittimfn.poi.enums;

import org.apache.poi.ss.usermodel.DateUtil;

import ittimfn.poi.App;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.math.BigDecimal;

public enum CellTypeEnum {
    NUMERIC(CellType.NUMERIC) { 
        @Override
        public String getCellValue(Cell cell) {
            return 
                DateUtil.isCellDateFormatted(cell) ?
                DATE.getCellValue(cell) : 
                BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString();
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            double fromValue = from.getNumericCellValue();
            to.setCellValue(fromValue);
        }
    },
    STRING(CellType.STRING) {
        @Override
        public String getCellValue(Cell cell) {
            return cell.getStringCellValue();
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            String fromValue = from.getStringCellValue();
            to.setCellValue(fromValue);
        }
    },
    FORMULA(CellType.FORMULA) {
        @Override
        public String getCellValue(Cell cell) {
            return cell.getCellFormula();
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            String fromValue = from.getCellFormula();
            to.setCellFormula(fromValue);
        }
    },
    BLANK(CellType.BLANK) {
        @Override
        public String getCellValue(Cell cell) {
            // TODO 機会があれば書く
            return null;
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            // TODO 暫定でString扱い。
            String fromValue = from.getStringCellValue();
            to.setCellValue(fromValue);
        }

    },
    BOOLEAN(CellType.BOOLEAN) {
        @Override
        public String getCellValue(Cell cell) {
            return Boolean.toString(cell.getBooleanCellValue());
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            boolean fromValue = from.getBooleanCellValue();
            to.setCellValue(fromValue);
        }
    },
    ERROR(CellType.ERROR) {
        @Override
        public String getCellValue(Cell cell) {
            // TODO 機会があれば書く
            return null;
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            byte fromValue = from.getErrorCellValue();
            to.setCellErrorValue(fromValue);
        }
    },
    DATE(null) {
        @Override
        public String getCellValue(Cell cell) {
            // TODO 機会があれば書く
            // TODO Date -> String変換については形式指定が必要。
            return null;
        }
        @Override
        public void copyValue(Cell from, Cell to) {
            // Excel的に日付は数値。
            NUMERIC.copyValue(from, to);
        }
    };
    
    private static Logger logger = LogManager.getLogger(CellTypeEnum.class);
    private CellType cellType;
    
    private CellTypeEnum(CellType cellType) {
        this.cellType = cellType;
    }

    public static CellTypeEnum valueOfCellType(CellType cellType) throws IllegalArgumentException {
        for(CellTypeEnum e : values()) { 
            if(e.cellType == cellType) {
                return e;
            }
        }
        throw new IllegalArgumentException("Not Found : " + cellType);
    }

    public abstract String getCellValue(Cell cell);
    public void copyCell(Cell from, Cell to) {
        logger.info(this.name() + " : " + from + " -> " + to);
        this.copyValue(from, to);
        this.copyStyle(from, to);
    };

    protected abstract void copyValue(Cell from, Cell to);
    private void copyStyle(Cell from, Cell to) {
        CellStyle fromStyle = from.getCellStyle();
        to.setCellStyle(fromStyle);
    }
}