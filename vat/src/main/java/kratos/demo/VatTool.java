package kratos.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.JsonNode;

import kratos.demo.enums.InvoiceTypeCodeEnum;
import kratos.demo.enums.VoidMarkEnum;
import kratos.demo.utils.HttpRequestUtils;
import kratos.demo.utils.JsonMapper;

/**
 * Hello world!
 *
 * @author QQQ-G
 */
public class VatTool {

    private static final String OPEN_URL = "https://open.leshui365.com";

    private static final String VOID_MARK = "voidMark";

    private static final String UNKNOWN = "未知错误";

    private static final int CHECK_DATE_INDEX = 1;

    private static final int INVOICE_TYPE_CODE_INDEX = 2;

    /**
     * 发票号码
     */
    private static final int INVOICE_NUMBER_INDEX = 3;

    /**
     * 发票代码
     */
    private static final int INVOICE_CODE_INDEX = 4;

    private static final int SALES_NAME_INDEX = 5;

    private static final int SALES_TAX_PAYER_NUM_INDEX = 6;

    private static final int PURCHASER_NAME_INDEX = 7;

    private static final int BILL_TIME_INDEX = 8;

    private static final int TOTAL_TAX_SUM_INDEX = 9;

    private static final int INVOICE_AMOUNT_INDEX = 10;

    private static final int CHECK_CODE_INDEX = 11;

    private static final int INVOICE_REMARKS_INDEX = 12;

    /**
     * 验真结果列
     */
    private static final int VOID_MARK_INDEX = 13;

    private static final int VAT_DETAIL_INDEX = 14;

    public static void main(String[] args) {
        VatTool vatTool = new VatTool();
        String path = "./";
        vatTool.vat(path);
    }

    public void vat(String path) {
        // 1. 获取所有xlsx文件
        List<String> filePathList = getPathList(path);
        // 2. 获取token
        String token = getToken();
        // 3. 循环处理xlsx文件。
        for (String thisPath : filePathList) {
            try {
                vatFile(path, thisPath, token);
            } catch (Exception e) {
                System.out.println(thisPath + "文件vat失败，错误信息如下：");
                System.out.println(e.getMessage());
            }
        }
    }

    private void vatFile(String path, String thisPath, String token) {
        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(thisPath));
        } catch (Exception e) {
            System.out.println("读取excel失败，filePath = " + thisPath);
            return;
        }
        XSSFSheet sheet = workbook.getSheet("ocr结果");
        int lastRowNum = sheet.getLastRowNum();
        for (int index = 1; index <= lastRowNum; index++) {
            Row row = sheet.getRow(index);
            if (row != null) {
                try {
                    vatLine(row, index, token);
                } catch (Exception e) {
                    Cell cell13 = row.createCell(VOID_MARK_INDEX, 1);
                    cell13.setCellValue(UNKNOWN);
                    Cell cell14 = row.createCell(VAT_DETAIL_INDEX, 1);
                    cell14.setCellValue("这一行vat失败，错误 = " + e.getMessage());
                    System.out.println(thisPath + "文件的第" + index + "行vat失败，错误信息如下：");
                    System.out.println(e.getMessage());
                }
            }
        }
        // 7. 刷流
        try {
            String date = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss").format(new Date());
            OutputStream out = new FileOutputStream(path + "\\Vat_Result_" + date + ".xlsx");
            workbook.write(out);
            out.flush();
            out.close();
        } catch (Exception e) {
            System.out.println("刷流写EXCEL文件失败");
            System.out.println(e.getMessage());
        }
    }

    private void vatLine(Row row, int index, String token) {
        Cell cell13 = row.getCell(VOID_MARK_INDEX);
        if (cell13 != null) {
            String cellInfo = getCellInfo(cell13);
            if (VoidMarkEnum.VALID.getName().equals(cellInfo)) {
                System.out.println("第" + index + "行，验真结果显示正常，此行不参与验真。");
                return;
            } else if (VoidMarkEnum.NOT_JPG.getName().equals(cellInfo)) {
                System.out.println("第" + index + "行，非jpg（不是发票数据），此行不参与验真。");
                return;
            }
        } else {
            cell13 = row.createCell(VOID_MARK_INDEX, 1);
        }
        Map<String, String> parameter = getParameter(row, token);
        String result = getVatInfoByParam(parameter);
        JavaType mapType = JsonMapper.nonDefaultMapper().contructMapType(HashMap.class, String.class, String.class);
        Map<String, String> resultMap = JsonMapper.nonDefaultMapper().fromJson(result, mapType);
        if (resultMap != null && resultMap.size() > 0) {
            String rtnCode = "00";
            String resultCode = "1000";
            if (StringUtils.equals(rtnCode, resultMap.get("RtnCode"))
                    && StringUtils.equals(resultCode, resultMap.get("resultCode"))) {
                String invoiceResult = resultMap.get("invoiceResult");
                JsonNode invoiceResultNode = JsonMapper.nonDefaultMapper().fromJson(invoiceResult, JsonNode.class);
                if (VoidMarkEnum.INVALID.getValue().equals(invoiceResultNode.get(VOID_MARK).asText())) {
                    cell13.setCellValue(VoidMarkEnum.INVALID.name());
                } else if (VoidMarkEnum.VALID.getValue().equals(invoiceResultNode.get(VOID_MARK).asText())) {
                    cell13.setCellValue(VoidMarkEnum.VALID.name());
                    setRowInfo(row, invoiceResultNode);
                } else {
                    cell13.setCellValue(UNKNOWN);
                }
            } else {
                String resultMsg = resultMap.get("resultMsg");
                cell13.setCellValue(resultMsg);
            }
        } else {
            cell13.setCellValue(UNKNOWN);
        }
        Cell cell14 = row.createCell(VAT_DETAIL_INDEX, 1);
        cell14.setCellValue(result);
    }

    private List<String> getPathList(String path) {
        File file = new File(path);
        List<String> filePathList = getFilePath(file);
        final String xlsx = "xlsx";
        return filePathList.stream().filter(thisPath -> {
            if (thisPath.contains(xlsx)) {
                return true;
            } else {
                System.out.println(thisPath + " 不是xlsx文件。");
                return false;
            }
        }).collect(Collectors.toList());
    }

    private Map<String, String> getParameter(Row row, String token) {
        // 发票号码
        String invoiceNumber = getCellInfo(row.getCell(INVOICE_NUMBER_INDEX));
        // 发票代码
        String invoiceCode = getCellInfo(row.getCell(INVOICE_CODE_INDEX));
        String billTime = getCellInfo(row.getCell(BILL_TIME_INDEX));
        billTime = billTime.replace("年", "-").replace("月", "-").replace("日", "");
        String invoiceAmount = getCellInfo(row.getCell(INVOICE_AMOUNT_INDEX));
        Map<String, String> map = new HashMap<>(16);
        map.put("invoiceNumber", invoiceNumber);
        map.put("invoiceCode", invoiceCode);
        map.put("billTime", billTime);
        map.put("invoiceAmount", invoiceAmount);
        String checkCode = getCellInfo(row.getCell(CHECK_CODE_INDEX));
        if (StringUtils.isNotBlank(checkCode)) {
            checkCode = checkCode.substring(checkCode.length() - 6);
            map.put("checkCode", checkCode);
        }
        map.put("token", token);
        return map;
    }

    public String getVatInfoByParam(Map<String, String> parameter) {
        String url = OPEN_URL + "/api/invoiceInfoForCom";
        String requestJson = JsonMapper.nonDefaultMapper().toJson(parameter);
        return HttpRequestUtils.sendPost(url, requestJson);
    }

    private void setRowInfo(Row row, JsonNode vatInfoNode) {
        Cell checkDate = getOrCreateCell(row, CHECK_DATE_INDEX);
        Cell invoiceTypeCode = getOrCreateCell(row, INVOICE_TYPE_CODE_INDEX);
        Cell salesName = getOrCreateCell(row, SALES_NAME_INDEX);
        Cell salesTaxpayerNum = getOrCreateCell(row, SALES_TAX_PAYER_NUM_INDEX);
        Cell purchaserName = getOrCreateCell(row, PURCHASER_NAME_INDEX);
        Cell totalTaxSum = getOrCreateCell(row, TOTAL_TAX_SUM_INDEX);
        Cell checkCode = getOrCreateCell(row, CHECK_CODE_INDEX);
        Cell invoiceRemarks = getOrCreateCell(row, INVOICE_REMARKS_INDEX);
        Cell voidMark = getOrCreateCell(row, VOID_MARK_INDEX);
        checkDate.setCellValue(vatInfoNode.get("checkDate").asText());
        invoiceTypeCode.setCellValue(InvoiceTypeCodeEnum.getName(vatInfoNode.get("invoiceTypeCode").asText()));
        salesName.setCellValue(vatInfoNode.get("salesName").asText());
        salesTaxpayerNum.setCellValue(vatInfoNode.get("salesTaxpayerNum").asText());
        purchaserName.setCellValue(vatInfoNode.get("purchaserName").asText());
        totalTaxSum.setCellValue(vatInfoNode.get("totalTaxSum").asText());
        checkCode.setCellValue(vatInfoNode.get("checkCode").asText());
        invoiceRemarks.setCellValue(vatInfoNode.get("invoiceRemarks").asText());
        voidMark.setCellValue(VoidMarkEnum.getName(vatInfoNode.get("voidMark").asText()));
    }

    private Cell getOrCreateCell(Row row, int index) {
        Cell cell = row.getCell(index);
        if (cell == null) {
            return row.createCell(index, 1);
        } else {
            return cell;
        }
    }

    public String getToken() {
        final String appKey = "";
        final String appSecret = "";
        String tokenRest = HttpRequestUtils.sendGet(OPEN_URL + "/getToken?appKey=" + appKey + "&appSecret=" + appSecret);
        try {
            JavaType mapType = JsonMapper.nonDefaultMapper().contructMapType(HashMap.class, String.class, String.class);
            Map<String, String> tokenMap = JsonMapper.nonDefaultMapper().fromJson(tokenRest, mapType);
            return tokenMap.get("token");
        } catch (Exception e) {
            System.out.println(e.getMessage());
            throw new RuntimeException("获取token失败");
        }
    }

    private String getCellInfo(Cell cell) {
        if (cell == null) {
            throw new RuntimeException("excel获取单元格数据失败，单元格是空");
        }
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            default:
                throw new RuntimeException("excel获取单元格数据失败，数据非字符串类型/非数字类型");
        }
    }

    private List<String> getFilePath(File file) {
        List<String> filePathList = new ArrayList<>();
        boolean directory = file.isDirectory();
        if (directory) {
            for (File listFile : Objects.requireNonNull(file.listFiles())) {
                List<String> subFilePathList = getFilePath(listFile);
                filePathList.addAll(subFilePathList);
            }
        } else {
            filePathList.add(file.getAbsolutePath());
        }
        return filePathList;
    }
}
