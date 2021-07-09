package XWPFDoc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class XConvertUtil {

    public static void convertDoc(String templatePath, Map<String, Object> content, String outputPath) throws Exception {
        FileInputStream fileInputStream = new FileInputStream(templatePath);
        //添加中转文件
        String tempFilePath = "src/main/resources/$temp$.docx";
        OutputStream outputFile = new FileOutputStream(outputPath);
        OutputStream tempOut = new FileOutputStream(tempFilePath);
        FileInputStream tempInput = new FileInputStream(tempFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);
        //初始化读取行数
        int rowNum = 1;
        //获取所有段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        Set<String> contentKeys = content.keySet();
        //开始替换文本操作
        for (XWPFParagraph p : paragraphs) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null && runs.size() > 0) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    //检查是否有要替换的内容，有的话则替换
                    for (String contentKey : contentKeys) {
                        //替换文本
                        if (content.get(contentKey) instanceof String) {
                            if (text != null && text.contains("${" + contentKey + "}")) {
                                text = text.replace("${" + contentKey + "}", (String) content.get(contentKey));
                            }
                        }
                    }
                    r.setText(text, 0);
                }
            }
        }
        //获取第一张表格
        XWPFTable xwpfTable = document.getTables().get(0);
        //开始表格中的替换文本操作
        for (XWPFTableRow row : xwpfTable.getRows()) {
            for (int i = 0; i < xwpfTable.getRow(rowNum).getTableCells().size(); i++) {
                for (String contentKey : contentKeys) {
                    if (content.get(contentKey) instanceof String) {
                        if (row.getCell(i).getText() != null && row.getCell(i).getText().contains("${" + contentKey + "}")) {
                            //晴空单元格内所有数据
                            for (XWPFRun run : row.getCell(i).getParagraphs().get(0).getRuns()) {
                                run.setText("", 0);
                            }
                            XWPFRun xwpfRun = row.getCell(i).getParagraphs().get(0).getRuns().get(0);
                            String text = (String) content.get(contentKey);
                            xwpfRun.setText(text, 0);
                        }
                    }
                }
            }
        }

        //获取替换内容
        List<String[]> tableContents = (List<String[]>) content.get("tableContent");
        if (tableContents != null && tableContents.size() > 0) {
            //新增一定行数
            if (tableContents.size() > 3) {
                for (int i = 0; i < (tableContents.size() - 3); i++) {
                    xwpfTable.addRow(xwpfTable.getRow(rowNum),1);
                }
                document.write(tempOut);

                document = new XWPFDocument(tempInput);
                xwpfTable = document.getTables().get(0);
            }

            //新增对应行数记录
            for (String[] tableContent : tableContents) {
                for (int i = 0; i < xwpfTable.getRow(rowNum).getTableCells().size(); i++) {
                    xwpfTable.getRow(rowNum).getCell(i).setText(tableContent[i]);
                }
                rowNum ++;
            }
        }
        //输出文件
        document.write(outputFile);
        tempInput.close();
        fileInputStream.close();
        //删除临时文件
        File f = new File(tempFilePath);
        if(f.exists()){
            f.delete();
        }
    }
}
