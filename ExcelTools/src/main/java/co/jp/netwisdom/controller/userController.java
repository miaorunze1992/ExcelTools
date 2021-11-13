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
import java.util.Arrays;
import java.util.List;

import com.alibaba.fastjson.*;


@RestController
@RequestMapping("/document")
public class userController {

    String str = "{\n" +
            "  \"templateId\": \"001\",\n" +
            "  \"userInfo\": {\n" +
            "    \"visaNumber\": \"EG82142208EA\",\n" +
            "    \"nationality\": \"1\",\n" +
            "    \"name\": \"HU XIN\",\n" +
            "    \"birthYear\": \"1996\",\n" +
            "    \"birthMonth\": \"01\",\n" +
            "    \"birthDay\": \"09\",\n" +
            "    \"visaType\": \"4\",\n" +
            "    \"visaMonths\": \"36\",\n" +
            "    \"visaExpireYear\":\"2024\",\n" +
            "    \"visaExpireMonth\": \"09\",\n" +
            "    \"visaExpireDay\": \"15\",\n" +
            "    \"visaExpireDate\": \"2024-09-15\",\n" +
            "    \"tel\": \"080-9402-8668\",\n" +
            "    \"address\": \"横滨市鶴見区尻手2-1-55グランドステッジ鶴見204\",\n" +
            "    \"graduateSchool\": \"東海大学\",\n" +
            "    \"graduateYear\": \"2021\",\n" +
            "    \"graduateMonth\": \"04\",\n" +
            "    \"passportNumber\": \"E39598799\",\n" +
            "    \"passportExpireYear\": \"2025\",\n" +
            "    \"passportExpireMonth\": \"03\",\n" +
            "    \"passportExpireDay\": \"25\",\n" +
            "    \"studyCondition\": \"1\"\n" +
            "  }\n" +
            "}\n";

    private static final String NW_EXCEL_001 = "在留资格変更申請"; // 在留资格変更申請
    private static final String NW_EXCEL_002 = "在留资格更新申請"; // 在留资格更新申請
    private static final String NW_EXCEL_003 = "労働条件通知書"; // 労働条件通知書
    private static final String NW_WORD_004 = "開発担当業務"; // 開発担当業務
    private static final String NW_WORD_005 = "雇用理由書"; // 雇用理由書

    /**
     * 选择下载模板
     */
    @GetMapping("/downloadDocument")
    public void downloadTemplate(HttpServletResponse response) throws Exception {
        // 测试用,正常jsonObj从前台传入
        JSONObject jsonObj = JSONObject.parseObject(str);
        // 资源路径
        String resourcePath;
        JSONObject ob = jsonObj.getJSONObject("userInfo");
        // 根据前台templateId进行区分下载模板
        switch (jsonObj.getString("templateId")) {
            case "001":
                resourcePath = String.format("excelTemplate/%s.xlsx", "001");
                downloadExcelFile(response, resourcePath, NW_EXCEL_001, ob);
                break;
            case "002":
                resourcePath = String.format("excelTemplate/%s.xlsx", "002");
                downloadExcelFile(response, resourcePath, NW_EXCEL_002, ob);
                break;
            case "003":
                resourcePath = String.format("excelTemplate/%s.xlsx", "003");
                downloadExcelFile(response, resourcePath, NW_EXCEL_003, ob);
                break;
            case "004":
                resourcePath = String.format("excelTemplate/%s.xlsx", "004");
                downloadExcelFile(response, resourcePath, NW_WORD_004, ob);
                break;
            case "005":
                resourcePath = String.format("excelTemplate/%s.xlsx", "005");
                downloadExcelFile(response, resourcePath, NW_WORD_005, ob);
                break;
            default:
                break;

        }
    }

    /**
     * @param response     响应
     * @param resourcePath excel模板路径
     */
    private static void downloadExcelFile(HttpServletResponse response, String resourcePath, String excelName, JSONObject jsonObject) throws IOException {

        InputStream inputStream = new ClassPathResource(resourcePath).getInputStream();
        Workbook workbook = new XSSFWorkbook(inputStream);

        Sheet sheet;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            // 获取Sheet
            sheet = workbook.getSheetAt(i);
            for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
                // 获取Row
                Row row = sheet.getRow(j);
                for (int k = 0; k < row.getLastCellNum(); k++) {
                    // 获取Cell
                    Cell cell = row.getCell(k);
                    if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
                        changeCellCount(cell, jsonObject);
                    }
                }
            }
        }
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("content-Disposition", "attachment;filename=" + URLEncoder.encode(excelName + ".xlsx", "utf-8"));
        OutputStream outputStream = response.getOutputStream();

        workbook.write(outputStream);
        // 关闭流
        outputStream.flush();
        outputStream.close();

    }

    /**
     * @param cell       单元格对象
     * @param jsonObject json对象
     */
    public static void changeCellCount(Cell cell, JSONObject jsonObject) {
        List<String> perchList = Arrays.asList(
                "${visaNumber}",    // パスポートナンバー
                "${nationality}",   // 国籍
                "${name}",  // 氏名
                "${birthYear}", // 出生月日
                "${birthMonth}",
                "${birthDay}",
                "${visaType}",  // 現に有する在留資格
                "${visaMonths}",    // 在留期限
                "${visaExpireYear}",    // 在留期間の満了日
                "${visaExpireMonth}",
                "${visaExpireDay}",
                "${tel}",   // 携帯電話
                "${address}",   // 住居地
                "${graduateSchool}",    // 学校名
                "${graduateYear}",  // 卒業年月
                "${graduateMonth}",
                "${passportNumber}",    // パスポートナンバー
                "${passportExpireYear}",    // 有効期限
                "${passportExpireMonth}",
                "${passportExpireDay}");

        for (String str : perchList) {
            if (cell.getStringCellValue().equals(str)) {
                cell.setCellValue(jsonObject.getString(str
                        .replace("${", "")
                        .replace("}", "")));
            }
        }
    }
}
