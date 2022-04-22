package kratos.demo.enums;

/**
 * Created on 2022/4/19.
 *
 * @author zhiqiang bao
 */
public enum VoidMarkEnum {

    /**
     * 正常
     */
    VALID("正常", "0"),
    /**
     * 作废
     */
    INVALID("作废", "1"),
    /**
     * 作废
     */
    NOT_JPG("文件非jpg", "9999");

    private final String name;

    private final String value;

    VoidMarkEnum(String name, String value) {
        this.name = name;
        this.value = value;
    }

    public static String getName(String value) {
        if (value.equals(VALID.value)) {
            return VALID.name;
        } else if (value.equals(INVALID.value)) {
            return INVALID.name;
        }
        return "未知结果";
    }

    public final String getValue() {
        return value;
    }

    public final String getName() {
        return name;
    }
}
