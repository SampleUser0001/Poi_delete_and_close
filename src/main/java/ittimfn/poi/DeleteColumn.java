package ittimfn.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ittimfn.poi.enums.CellTypeEnum;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class DeleteColumn {

    private static Logger logger = LogManager.getLogger(DeleteColumn.class);

    public void delete(String input, String output) throws IOException {
        logger.info("input: " + input);
        FileInputStream fis = new FileInputStream(input);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        int columnToDelete = 1; // 削除したい列のインデックス (0から始まる)

        this.deleteColumns(sheet, columnToDelete);

        fis.close();

        logger.info("output: " + output);
        FileOutputStream fos = new FileOutputStream(output);
        workbook.write(fos);
        fos.close();
        workbook.close();
    }

    private void deleteColumns(Sheet sheet, int columnToDelete) {
        for (Row row : sheet) {
            for (int colIndex = columnToDelete; colIndex < row.getLastCellNum() - 1; colIndex++) {
                logger.trace("row: {}, colIndex: {}", row.getRowNum(), colIndex);

                Cell oldCell = row.getCell(colIndex);
                logger.trace("oldCell: {}", oldCell);

                Cell newCell = row.getCell(colIndex + 1);
                logger.trace("newCell: {}", newCell);

                if (oldCell == null && newCell == null) {
                    // 両方ともない。何もしない。
                } else if (oldCell == null && newCell != null) {
                    // 削除対象セルには何もないが、右に値がある
                    logger.trace("newCell.getCellType(): {}", newCell.getCellType());
                    logger.trace("newCell value: {}", CellTypeEnum.valueOfCellType(newCell.getCellType()).getCellValue(newCell));

                    oldCell = row.createCell(colIndex);
                    oldCell.setCellType(newCell.getCellType());
                    oldCell.setCellValue(CellTypeEnum.valueOfCellType(newCell.getCellType()).getCellValue(newCell));

                } else if (oldCell != null && newCell == null) {
                    // 削除対象セルにはなにか書かれているが、右に値がない
                    logger.trace("oldCell.getCellType(): {}", oldCell.getCellType());
                    logger.trace("oldCell value: {}", CellTypeEnum.valueOfCellType(oldCell.getCellType()).getCellValue(oldCell));

                    row.removeCell(oldCell);
                } else {
                    logger.trace("oldCell.getCellType(): {}", oldCell.getCellType());
                    logger.trace("oldCell value: {}", CellTypeEnum.valueOfCellType(oldCell.getCellType()).getCellValue(oldCell));
                    logger.trace("newCell.getCellType(): {}", newCell.getCellType());
                    logger.trace("newCell value: {}", CellTypeEnum.valueOfCellType(newCell.getCellType()).getCellValue(newCell));

                    // 両方のセルに値がある
                    oldCell.setCellType(newCell.getCellType());
                    oldCell.setCellValue(CellTypeEnum.valueOfCellType(newCell.getCellType()).getCellValue(newCell));
                }

            }
            // 最後の列を削除
            Cell lastCell = row.getCell(row.getLastCellNum() - 1);
            if (lastCell != null) {
                row.removeCell(lastCell);
            }
        }
    }
}