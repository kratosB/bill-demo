package kratos.demo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.JsonNode;
import com.tencentcloudapi.common.Credential;
import com.tencentcloudapi.common.profile.ClientProfile;
import com.tencentcloudapi.common.profile.HttpProfile;
import com.tencentcloudapi.ocr.v20181119.OcrClient;
import com.tencentcloudapi.ocr.v20181119.models.VatInvoiceOCRRequest;
import com.tencentcloudapi.ocr.v20181119.models.VatInvoiceOCRResponse;

import net.coobird.thumbnailator.Thumbnails;
import sun.misc.BASE64Encoder;

/**
 * Created on 2022/4/12.
 *
 * @author zhiqiang bao
 */
public class OcrTool {

    public static void main(String[] args) {
        String path = "./";
        OcrTool ocrTool = new OcrTool();
        ocrTool.ocrVerify(path);
    }

    public void ocrVerify(String path) {
        StringBuilder errorMessage = new StringBuilder();
        // 1. 读所有文件路径
        List<String> filePathList = getFilePath(new File(path));
        // 2. 创建excel
        Workbook workBook = new XSSFWorkbook();
        Sheet sheet = workBook.createSheet("ocr结果");
        initExcel(sheet.createRow(0));
        // 文件容量限制，如果大于这个，需要压缩
        int maxlength = 7396760;
        for (int i = 0; i < filePathList.size(); i++) {
            String filePath = filePathList.get(i);
            String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);
            if (!filePath.contains("jpg") && !filePath.contains("jpeg")) {
                Row row = sheet.createRow(i + 1);
                writeBlankRow(fileName, row, "文件非jpg");
                continue;
            }
            // 3. 获取base64
            try {
                String encodedValue = getEncodedValue(filePath, maxlength);
                try {
                    // 4. 发票识别，数据解析，文件写入
                    String ocrValue = ocr(encodedValue);
                    System.out.println(fileName + " = " + ocrValue);
                    // 5. 提取数据
                    HashMap<Integer, String> map = getMap(ocrValue);
                    // 6. 写excel
                    Row row = sheet.createRow(i + 1);
                    writeRow(map, fileName, row);
                } catch (Exception e) {
                    errorMessage.append("发生错误，fileName = ").append(fileName).append(", encodeValue.length = ")
                            .append(encodedValue.length()).append(", e = ").append(e.getMessage()).append("\n");
                    Row row = sheet.createRow(i + 1);
                    Cell cell0 = row.createCell(0, 1);
                    cell0.setCellValue(fileName);
                }
            } catch (Exception e) {
                errorMessage.append("文件名为《").append(fileName).append("》的文件读取错误，错误 = （").append(e.getMessage()).append(")")
                        .append("\n");
            }
        }
        // 7. 刷流
        try {
            String date = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss").format(new Date());
            OutputStream out = new FileOutputStream(path + "\\Ocr_Result_" + date + ".xlsx");
            workBook.write(out);
            out.flush();
            out.close();
        } catch (Exception e) {
            errorMessage.append("错误 = （").append(e.getMessage()).append(")");
        }
        System.out.println(errorMessage);
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

    private String getEncodedValue(String filePath, int maxLength) throws Exception {
        InputStream fileInputStream = new FileInputStream(filePath);
        int available = fileInputStream.available();
        byte[] bytes = new byte[available];
        fileInputStream.read(bytes);
        fileInputStream.close();
        return encode(bytes, maxLength);
    }

    private String encode(byte[] bytes, int maxLength) throws Exception {
        String encode = new BASE64Encoder().encode(bytes);
        if (encode.length() > maxLength) {
            InputStream inputStream = new ByteArrayInputStream(bytes);
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            Thumbnails.of(inputStream).scale(1f).outputQuality(0.95f).toOutputStream(outputStream);
            byte[] newBytes = outputStream.toByteArray();
            return encode(newBytes, maxLength);
        } else {
            return encode;
        }
    }

    private String ocr(String encodeValue) throws Exception {
        Credential cred = new Credential("secretId", "secretKey");
        HttpProfile httpProfile = new HttpProfile();
        httpProfile.setEndpoint("ocr.ap-shanghai.tencentcloudapi.com");
        ClientProfile clientProfile = new ClientProfile();
        clientProfile.setHttpProfile(httpProfile);
        OcrClient client = new OcrClient(cred, "ap-shanghai", clientProfile);
        String params = "{\"ImageBase64\":\"" + encodeValue + "\"}";
        VatInvoiceOCRRequest req = VatInvoiceOCRRequest.fromJsonString(params, VatInvoiceOCRRequest.class);
        VatInvoiceOCRResponse resp = client.VatInvoiceOCR(req);
        return VatInvoiceOCRRequest.toJsonString(resp);
    }

    private HashMap<Integer, String> getMap(String resultString) {
        JsonNode jsonNode = JsonMapper.nonEmptyMapper().fromJson(resultString, JsonNode.class);
        JsonNode vatInvoiceInfos = jsonNode.path("VatInvoiceInfos");
        HashMap<Integer, String> result = new HashMap<>(16);
        String number = "发票号码";
        String serialNo = "发票代码";
        String sellerName = "销售方名称";
        String buyerName = "购买方名称";
        String date = "开票日期";
        String amount = "小写金额";
        String amountWithoutTax = "合计金额";
        String checkNo = "校验码";
        String comment = "备注";
        vatInvoiceInfos.forEach(node -> {
            String name = node.get("Name").asText();
            if (StringUtils.equals(name, number)) {
                result.put(NUMBER_INDEX, node.get("Value").asText().replace("No", ""));
            } else if (StringUtils.equals(name, serialNo)) {
                result.put(SERIAL_NO_INDEX, node.get("Value").asText());
            } else if (StringUtils.equals(name, sellerName)) {
                result.put(SELLER_NAME_INDEX, node.get("Value").asText());
            } else if (StringUtils.equals(name, buyerName)) {
                result.put(BUYER_NAME_INDEX, node.get("Value").asText());
            } else if (StringUtils.equals(name, date)) {
                result.put(DATE_INDEX, node.get("Value").asText());
            } else if (StringUtils.equals(name, amount)) {
                result.put(AMOUNT_INDEX, node.get("Value").asText().replace("¥", ""));
            } else if (StringUtils.equals(name, amountWithoutTax)) {
                result.put(AMOUNT_WITHOUT_TAX_INDEX, node.get("Value").asText().replace("¥", ""));
            } else if (StringUtils.equals(name, checkNo)) {
                result.put(CHECK_NO_INDEX, node.get("Value").asText());
            } else if (StringUtils.equals(name, comment)) {
                result.put(COMMENT_INDEX, node.get("Value").asText());
            }
        });
        return result;
    }

    private void writeBlankRow(String fileName, Row row, String message) {
        Cell cell0 = row.createCell(FILE_NAME_INDEX, 1);
        cell0.setCellValue(fileName);
        Cell cell9 = row.createCell(VOID_MARK_INDEX, 1);
        cell9.setCellValue(message);
    }

    private void writeRow(HashMap<Integer, String> map, String fileName, Row row) {
        Cell cell0 = row.createCell(FILE_NAME_INDEX, 1);
        cell0.setCellValue(fileName);
        if (map != null) {
            map.forEach((index, value) -> {
                Cell cell = row.createCell(index, 1);
                cell.setCellValue(value);
            });
        }
    }

    private static final int FILE_NAME_INDEX = 0;

    /**
     * 发票号码
     */
    private static final int NUMBER_INDEX = 3;

    /**
     * 发票代码
     */
    private static final int SERIAL_NO_INDEX = 4;

    private static final int SELLER_NAME_INDEX = 5;

    private static final int BUYER_NAME_INDEX = 7;

    private static final int DATE_INDEX = 8;

    private static final int AMOUNT_INDEX = 9;

    private static final int AMOUNT_WITHOUT_TAX_INDEX = 10;

    private static final int CHECK_NO_INDEX = 11;

    private static final int COMMENT_INDEX = 12;

    private static final int VOID_MARK_INDEX = 13;

    private static final int VAT_DETAIL_INDEX = 14;

    private void initExcel(Row row) {
        Cell fileName = row.createCell(FILE_NAME_INDEX, 1);
        Cell checkDate = row.createCell(1, 1);
        Cell invoiceTypeCode = row.createCell(2, 1);
        Cell number = row.createCell(NUMBER_INDEX, 1);
        Cell serialNo = row.createCell(SERIAL_NO_INDEX, 1);
        Cell sellerName = row.createCell(SELLER_NAME_INDEX, 1);
        Cell salesTaxpayerNum = row.createCell(6, 1);
        Cell buyerName = row.createCell(BUYER_NAME_INDEX, 1);
        Cell date = row.createCell(DATE_INDEX, 1);
        Cell amount = row.createCell(AMOUNT_INDEX, 1);
        Cell amountWithoutTax = row.createCell(AMOUNT_WITHOUT_TAX_INDEX, 1);
        Cell checkNo = row.createCell(CHECK_NO_INDEX, 1);
        Cell comment = row.createCell(COMMENT_INDEX, 1);
        Cell voidMark = row.createCell(VOID_MARK_INDEX, 1);
        Cell detail = row.createCell(VAT_DETAIL_INDEX, 1);
        fileName.setCellValue("文件名称");
        checkDate.setCellValue("查询时间");
        invoiceTypeCode.setCellValue("票种");
        number.setCellValue("发票号码");
        serialNo.setCellValue("发票代码");
        sellerName.setCellValue("销售方名称");
        salesTaxpayerNum.setCellValue("销售方纳税人识别号");
        buyerName.setCellValue("购买方名称");
        date.setCellValue("开票日期");
        amount.setCellValue("小写金额");
        amountWithoutTax.setCellValue("不含税金额");
        checkNo.setCellValue("校验码");
        comment.setCellValue("备注");
        voidMark.setCellValue("验真结果");
        detail.setCellValue("发票详情");
    }
}
