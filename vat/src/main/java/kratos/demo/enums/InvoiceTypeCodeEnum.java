package kratos.demo.enums;

/**
 * Created on 2022/4/19.
 *
 * @author zhiqiang bao
 */
public enum InvoiceTypeCodeEnum {

    /**
     * 增值税专用发票
     */
    ONE("增值税专用发票", "01"),
    /**
     * 机动车销售统一发票
     */
    THREEE("机动车销售统一发票", "03"),
    /**
     * 增值税普通发票
     */
    FOUR("增值税普通发票", "04"),
    /**
     * 增值税电子专用发票
     */
    EIGHT("增值税电子专用发票", "08"),
    /**
     * 增值税电子普通发票
     */
    TEN("增值税电子普通发票", "10"),
    /**
     *
     */
    ELEVEN("卷式普通发票", "11");

    private final String name;

    private final String value;

    InvoiceTypeCodeEnum(String name, String value) {
        this.name = name;
        this.value = value;
    }

    public static String getName(String value) {
        for (InvoiceTypeCodeEnum invoiceTypeCodeEnum : InvoiceTypeCodeEnum.values()) {
            if (invoiceTypeCodeEnum.value.equals(value)) {
                return invoiceTypeCodeEnum.name;
            }
        }
        return "发票类型代码为" + value + ", 没找到对应翻译。";
    }
}
