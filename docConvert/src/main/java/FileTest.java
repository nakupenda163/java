import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class FileTest {
    public static void main(String[] args) throws IOException {
        String abc = "你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好" +
                "你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好你好" +
                "你好你好你好你好你好你好你好你好你好你好你好";
        String filePath = "src/main/resources/download.docx";
        FileOutputStream fos = new FileOutputStream(filePath);
        fos.write(abc.getBytes("gbk"));
        fos.close();
    }
}
