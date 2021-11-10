package co.jp.netwisdom.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
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

        if (templateId == null) {
            return;
        }
        String resourcePath;
        switch (templateId) {
            case "001":
                resourcePath = String.format("excelTemplate/%s.xlsx", "001");
                downloadFile(response, resourcePath, NW_EXCEL_001);
                break;
            case "002":
                resourcePath = String.format("excelTemplate/%s.xlsx", "002");
                downloadFile(response, resourcePath, NW_EXCEL_002);
                break;
            case "003":
                resourcePath = String.format("excelTemplate/%s.xlsx", "003");
                downloadFile(response, resourcePath, NW_EXCEL_003);
                break;
            case "004":
                resourcePath = String.format("excelTemplate/%s.xlsx", "004");
                downloadFile(response, resourcePath, NW_WORD_004);
                break;
            case "005":
                resourcePath = String.format("excelTemplate/%s.xlsx", "005");
                downloadFile(response, resourcePath, NW_WORD_005);
                break;
            default:
                break;

        }


    }

    private static void downloadFile(HttpServletResponse response, String resourcePath, String fileName) throws IOException {
        ClassPathResource classPathResource = new ClassPathResource(resourcePath);
        InputStream inputStream = classPathResource.getInputStream();
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (Exception ex) {
            workbook = new HSSFWorkbook(inputStream);
        }
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName + ".xlsx", "utf-8"));
        response.setHeader("Access-Control-Expose-Headers", "content-Disposition");
        OutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();

    }

}
