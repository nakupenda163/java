package XWPFDoc;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ConvertTest {
    public static void main(String[] args) throws Exception {
        String templatePath = "src/main/resources/template2.docx";
        String outputPath = "src/main/resources/test2.docx";
        Map<String, Object> content = new HashMap<>();
        content.put("founder", "发起人");
        content.put("sa", "民间组织");
        content.put("accountName", "林锋");
        content.put("account", "12138");
        content.put("expireYear", "2021");
        content.put("expireMonth", "06");
        content.put("expireDay", "22");
        content.put("expireHour", "20");
        content.put("expireMin", "16");
        content.put("rmbAmt", "一百");
        content.put("rmbSmallAmt", "100");
        content.put("total", "500.00");
        Calendar instance = Calendar.getInstance();
        content.put("year", instance.get(Calendar.YEAR) + "");
        content.put("month", instance.get(Calendar.MONTH) + 1 + "");
        content.put("day", instance.get(Calendar.DAY_OF_MONTH) + "");

        String[] tableContent = new String[]{"1", "林峰", "2021-06-30", "rmb", "100.00", "摘要"};
        String[] tableContent2 = new String[]{"2", "林峰", "2021-06-30", "rmb", "100.00", "摘要"};
        String[] tableContent3 = new String[]{"3", "林峰", "2021-06-30", "rmb", "100.00", "摘要"};
        String[] tableContent4 = new String[]{"4", "林峰", "2021-06-30", "rmb", "100.00", "摘要"};
        String[] tableContent5 = new String[]{"5", "林峰", "2021-06-30", "rmb", "100.00", "摘要"};
        List<String[]> contents = new ArrayList<>();
        contents.add(tableContent);
        contents.add(tableContent2);
        contents.add(tableContent3);
        contents.add(tableContent4);
        contents.add(tableContent5);
        content.put("tableContent", contents);

        XConvertUtil.convertDoc(templatePath, content, outputPath);
    }
}
