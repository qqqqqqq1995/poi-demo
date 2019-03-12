import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author xuqiang
 * @date 2019/3/12
 */
public class POIUtil {

    public static Object getValue(Cell cell) {
        CellType cellType = cell.getCellType();
        if (cellType == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        }
        if (cellType == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        return cell.getStringCellValue();
    }

    public static List<List<List>> readExcel(String filePath) throws IOException {
        return readExcel(new File(filePath));
    }

    public static List<List<List>> readExcel(File file) throws IOException {
        Workbook workbook = WorkbookFactory.create(file);
        // excel 对象
        List<List<List>> excel = new ArrayList<>(workbook.getNumberOfSheets());
        List<List> sheetList;
        List rowList;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            // 保存sheet中的数据
            sheetList = new ArrayList<>();
            Sheet sheet = workbook.getSheetAt(i);
            for (int j = 0; j < sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                // 保存行中的数据
                rowList = new ArrayList(sheet.getLastRowNum());
                for (short k = 0; k < row.getLastCellNum(); k++) {
                    rowList.add(getValue(row.getCell(k)));
                }
                sheetList.add(rowList);
            }
            excel.add(sheetList);
        }
        return excel;
    }

    public static void main(String[] args) throws IOException {
        List<List<List>> lists = readExcel("C:\\Users\\15968\\Desktop\\工程导入模板.xls");
        System.out.println(lists.toString());
    }
}
