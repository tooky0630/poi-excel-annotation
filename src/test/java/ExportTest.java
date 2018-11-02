import com.penghaohuan.excel.handler.ExcelExporter;
import example.ExampleVO;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExportTest {

    private static final int EXPORT_SIZE = 10;

    @Test
    public void testImportExampleVO() throws IOException {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(new File("demo.xls"));
            ExcelExporter<ExampleVO> util = new ExcelExporter<>(ExampleVO.class);
            util.exportExcel(initExportList(), "Export Example", 60000, out);
            System.out.println("----执行完毕----------");
        } finally {
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private List<ExampleVO> initExportList() {
        final List<ExampleVO> list = new ArrayList<>(EXPORT_SIZE);
        for (int i = 0; i < EXPORT_SIZE; i++) {
            final ExampleVO row = new ExampleVO();
            row.setNo(i + "");
            row.setAge(10);
            row.setBirth(new Date());
            row.setPhone("13711111111");
            list.add(row);
        }
        return list;
    }
}
