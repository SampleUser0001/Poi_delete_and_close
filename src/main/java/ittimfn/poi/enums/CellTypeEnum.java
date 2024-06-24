package ittimfn.poi.enums;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Cell;
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
    },
    STRING(CellType.STRING) {
        @Override
        public String getCellValue(Cell cell) {
            return cell.getStringCellValue();
        }
    },
    FORMULA(CellType.FORMULA) {
        @Override
        public String getCellValue(Cell cell) {
            return cell.getCellFormula();
        }
    },
    BLANK(CellType.BLANK) {
        @Override
        public String getCellValue(Cell cell) {
            // TODO 機会があれば書く
            return null;
        }

    },
    BOOLEAN(CellType.BOOLEAN) {
        @Override
        public String getCellValue(Cell cell) {
            return Boolean.toString(cell.getBooleanCellValue());
        }
    },
    ERROR(CellType.ERROR) {
        @Override
        public String getCellValue(Cell cell) {
            // TODO 機会があれば書く
            return null;
        }
    },
    DATE(null) {
        @Override
        public String getCellValue(Cell cell) {
            // TODO 機会があれば書く
            // TODO Date -> String変換については形式指定が必要。
            return null;
        }
    };
    
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
}