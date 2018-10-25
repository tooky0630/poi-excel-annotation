import com.penghaohuan.excel.exception.ExcelTemplateException;
import com.penghaohuan.excel.exception.ExcelValidateException;
import com.penghaohuan.excel.handler.ExcelImporter;
import example.ExampleVO;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

public class ImportTest {

    private static final String FILE_NAME = "example.xlsx";

    @Test
    public void testImportExampleVO() throws FileNotFoundException, ExcelValidateException, ExcelTemplateException {
        ExcelImporter<ExampleVO> importer = new ExcelImporter<>(ExampleVO.class);
        final List<ExampleVO> exampleList = importer.importExcel(new FileInputStream(new File(FILE_NAME)), 1);
        assert exampleList != null;
    }
}
