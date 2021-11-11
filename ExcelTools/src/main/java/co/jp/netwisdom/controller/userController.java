package co.jp.netwisdom.controller;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

@RestController
@RequestMapping("/document")
public class userController {

    private static final String NW_EXCEL_001 = "在留资格変更申請"; // 在留资格変更申請
    private static final String NW_EXCEL_002 = "在留资格更新申請"; // 在留资格更新申請
    private static final String NW_EXCEL_003 = "労働条件通知書"; // 労働条件通知書
    private static final String NW_WORD_004 = "開発担当業務"; // 開発担当業務
    private static final String NW_WORD_005 = "雇用理由書"; // 雇用理由書

    /**
     * 选择下载模板
     */
    @GetMapping("/downloadDocument")
    public void downloadTemplate(HttpServletResponse response, String templateId) throws Exception {
        // templateId未指定
        if (templateId == null) {
            response.sendError(HttpServletResponse.SC_NOT_FOUND);
            return;
        }
        // 资源路径
        String resourcePath;
        // 根据前台templateId进行区分
        switch (templateId) {
            case "001":
                resourcePath = String.format("excelTemplate/%s.xlsx", "001");
                downloadExcelFile(response, resourcePath, NW_EXCEL_001);
                break;
            case "002":
                resourcePath = String.format("excelTemplate/%s.xlsx", "002");
                downloadExcelFile(response, resourcePath, NW_EXCEL_002);
                break;
            case "003":
                resourcePath = String.format("excelTemplate/%s.xlsx", "003");
                downloadExcelFile(response, resourcePath, NW_EXCEL_003);
                break;
            case "004":
                resourcePath = String.format("excelTemplate/%s.xlsx", "004");
                downloadExcelFile(response, resourcePath, NW_WORD_004);
                break;
            case "005":
                resourcePath = String.format("excelTemplate/%s.xlsx", "005");
                downloadExcelFile(response, resourcePath, NW_WORD_005);
                break;
            default:
                break;

        }


    }

    /**
     * @param response     响应
     * @param resourcePath excel模板路径
     */
    private static void downloadExcelFile(HttpServletResponse response, String resourcePath, String excelName) throws IOException {

        InputStream inputStream = new ClassPathResource(resourcePath).getInputStream();
        Workbook workbook = new XSSFWorkbook(inputStream);

        Sheet sheet;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            // 获取Sheel
            sheet = workbook.getSheetAt(i);
            for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
                // 获取Row
                Row row = sheet.getRow(j);
                for (int k = 0; k < row.getLastCellNum(); k++) {
                    // 获取Cell
                    Cell cell = row.getCell(k);
                    if (cell != null && cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().equals("${name}")) {
                        // 设置 Cell
                        // 根据前台传入数据，需要针对性替换
                        row.getCell(k).setCellValue("法外狂徒张三");
                    }
                }
            }
        }
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("content-Disposition", "attachment;excelName=" + URLEncoder.encode(excelName + ".xlsx", "utf-8"));
        OutputStream outputStream = response.getOutputStream();

        workbook.write(outputStream);
        // 关闭流
        outputStream.flush();
        outputStream.close();

    }

}
