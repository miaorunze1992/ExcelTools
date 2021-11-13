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
import java.util.HashMap;

import com.alibaba.fastjson.*;


@RestController
@RequestMapping("/document")
public class userController {

    String str = "{\n" +
            "  \"user\": {\n" +
            "    \"templateId\": \"001\",\n" +
            "    \"userInfo\": {\n" +
            "      \"visaNumber\": \"EG82142208EA\",\n" +
            "      \"nationality\": \"1\",\n" +
            "      \"name\": \"HU XIN\",\n" +
            "      \"birthday\": \"1996-01-09\",\n" +
            "      \"visaType\": \"4\",\n" +
            "      \"visaMonths\": \"36\",\n" +
            "      \"visaExpireDate\": \"2024-09-15\",\n" +
            "      \"tel\": \"080-9402-8668\",\n" +
            "      \"address\": \"横滨市鶴見区尻手2-1-55グランドステッジ鶴見204\",\n" +
            "      \"graduateSchool\": \"東海大学\",\n" +
            "      \"graduateDate\": \"2021-04\",\n" +
            "      \"passportNumber\": \"E39598799\",\n" +
            "      \"passportExpireDate\": \"2025-03-25\",\n" +
            "      \"studyCondition\": \"1\"\n" +
            "    }\n" +
            "  }\n" +
            "}";

    private static final String NW_EXCEL_001 = "在留资格変更申請"; // 在留资格変更申請
    private static final String NW_EXCEL_002 = "在留资格更新申請"; // 在留资格更新申請
    private static final String NW_EXCEL_003 = "労働条件通知書"; // 労働条件通知書
    private static final String NW_WORD_004 = "開発担当業務"; // 開発担当業務
    private static final String NW_WORD_005 = "雇用理由書"; // 雇用理由書

    /**
     * 选择下载模板
     */
    @GetMapping("/downloadDocument")
    //需要前端发来发来的json来进行操作，
    //本类因为模拟，在类中写了一个json文件，所以不用在此方法的参数中写 JSONObject jsonObject
    public void downloadTemplate(HttpServletResponse response) throws Exception {
        JSONObject jsonObject = JSONObject.parseObject(str);;
        //接收前端发来的JSON的话
        //需要写 jsonObject = jsonObject; 用本类中的JSON接收传来的JSON,这样方便，接下来方法的使用
        // templateId未指定
        JSONObject js = (JSONObject)jsonObject.get("user");
        String templateId = js.get("templateId").toString();
        System.out.println(templateId);
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
                downloadExcelFile(response, resourcePath, NW_EXCEL_001, (JSONObject) js.get("userInfo"));
                break;
            case "002":
                resourcePath = String.format("excelTemplate/%s.xlsx", "002");
                downloadExcelFile(response, resourcePath, NW_EXCEL_002, (JSONObject) js.get("userInfo"));
                break;
            case "003":
                resourcePath = String.format("excelTemplate/%s.xlsx", "003");
                downloadExcelFile(response, resourcePath, NW_EXCEL_003, (JSONObject) js.get("userInfo"));
                break;
            case "004":
                resourcePath = String.format("excelTemplate/%s.xlsx", "004");
                downloadExcelFile(response, resourcePath, NW_WORD_004, (JSONObject) js.get("userInfo"));
                break;
            case "005":
                resourcePath = String.format("excelTemplate/%s.xlsx", "005");
                downloadExcelFile(response, resourcePath, NW_WORD_005, (JSONObject) js.get("userInfo"));
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
                    changeCellCount(cell,jsonObject);
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

    //针对前台数据，需要针对性替换
    public static void changeCellCount(Cell cell, JSONObject jsonObject) {
        if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {

            //パスポートナンバー
            if (cell.getStringCellValue().equals("${visaNumber}")) {
                cell.setCellValue(jsonObject.getString("visaNumber"));
            }
            //国籍
            if (cell.getStringCellValue().equals("${nationality}")) {
                cell.setCellValue(jsonObject.getString("nationality"));
            }
            //氏名
            if (cell.getStringCellValue().equals("${name}")) {
                cell.setCellValue(jsonObject.getString("name"));
            }
            //出生月日
            String birthday = jsonObject.getString("birthday");
            if (cell.getStringCellValue().equals("${birthYear}")) {
                cell.setCellValue(birthday.substring(0, 4));
            }

            if (cell.getStringCellValue().equals("${birthMonth}")) {
                cell.setCellValue(birthday.substring(5, 7));
            }
            if (cell.getStringCellValue().equals("${birthDay}")) {
                cell.setCellValue(birthday.substring(8));
            }
            //現に有する在留資格
            if (cell.getStringCellValue().equals("${visaType}")) {
                cell.setCellValue(jsonObject.getString("visaType"));
            }
            //在留期限
            if (cell.getStringCellValue().equals("${visaMonths}")) {
                cell.setCellValue(jsonObject.getString("visaMonths"));
            }
            //在留期間の満了日
            String visaExpireDate = jsonObject.getString("visaExpireDate");
            if (cell.getStringCellValue().equals("${visaExpireYear}")) {
                cell.setCellValue(visaExpireDate.substring(0, 4));
            }
            if (cell.getStringCellValue().equals("${visaExpireMonth}")) {
                cell.setCellValue(visaExpireDate.substring(5, 7));
            }
            if (cell.getStringCellValue().equals("${visaExpireDay}")) {
                cell.setCellValue(visaExpireDate.substring(8));
            }
            //携帯電話
            if (cell.getStringCellValue().equals("${tel}")) {
                cell.setCellValue(jsonObject.getString("tel"));
            }
            //住居地
            if (cell.getStringCellValue().equals("${address}")) {
                cell.setCellValue(jsonObject.getString("address"));
            }
            //学校名
            if (cell.getStringCellValue().equals("${graduateSchool}")) {
                cell.setCellValue(jsonObject.getString("graduateSchool"));
            }
            //卒業年月
            String graduateDate = jsonObject.getString("graduateDate");
            if (cell.getStringCellValue().equals("${graduateYear}")) {
                cell.setCellValue(graduateDate.substring(0, 4));
            }
            if (cell.getStringCellValue().equals("${graduateMonth}")) {
                cell.setCellValue(graduateDate.substring(5));
            }
            //パスポートナンバー
            if (cell.getStringCellValue().equals("${passportNumber}")) {
                cell.setCellValue(jsonObject.getString("passportNumber"));
            }
            //有効期限
            String passportExpireDate = jsonObject.getString("passportExpireDate");
            if (cell.getStringCellValue().equals("${passportExpireYear}")) {
                cell.setCellValue(passportExpireDate.substring(0, 4));
            }
            if (cell.getStringCellValue().equals("${passportExpireMonth}")) {
                cell.setCellValue(passportExpireDate.substring(5, 7));
            }
            if (cell.getStringCellValue().equals("${passportExpireDay}")) {
                cell.setCellValue(passportExpireDate.substring(8));
            }


        }
    }
}
