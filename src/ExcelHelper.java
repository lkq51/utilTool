
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by lkq on 2017/4/10.
 */
public class ExcelHelper {
    public void handleCell() throws Exception {
        WritableCellFormat format = new WritableCellFormat();
        WritableSheet sheet = null;
        OutputStream outputStream = null;
        WritableWorkbook writableWorkbook = null;

        List<Object> excel2List = getExcleList("");
        outputStream = new FileOutputStream("");
        writableWorkbook = Workbook.createWorkbook(outputStream);
        sheet = writableWorkbook.createSheet("",1);
        for (int i = 0; i < excel2List.size(); i++){
            String keyWord = (String) excel2List.get(i);
            if (StringUtils.isNotBlank(keyWord)){
                //TODO
            }
            this.setCellData(sheet, new int[] { 0 }, i, keyWord, format, true);
            System.out.println("============" + excel2List.size());
            System.out.println(i);
        }
        writableWorkbook.write();
        writableWorkbook.close();
        outputStream.flush();
        outputStream.close();
    }

    private List<Object> getExcleList(String path){
        List<Object> excelList = new ArrayList<Object>();
        //创建单元格对象
        Cell cell  = null;
        try {
            //获取excel文件
            Workbook excel = Workbook.getWorkbook(new File(path));
            //获取第一个工作表对象
            Sheet sheet = excel.getSheet(0);

            for (int i = 0; i < sheet.getRows(); i++){
                String keyWord = null;
                for (int j = 0; j < sheet.getColumns(); j++){
                    cell = sheet.getCell(0, i);
                    keyWord = StringUtils.trim(cell.getContents());
                }
                excelList.add(keyWord);
            }
            excel.close();
        }catch (Exception e){
            e.printStackTrace();
        }
        return excelList;
    }

    private int setCellData(WritableSheet sheet, int[] cols, int rows, String strText, CellFormat format, boolean boolean_InsertRow) throws Exception{
        if (null == strText) {
            sheet.addCell(new Label(cols[0], rows, "", format));
            return 1;
        }
        int intMaxBytesPerline = 0;//一行能占用的最大字节数
        for (int i = 0; i < cols.length; i++)
            /*列之间的间隔？*/
            intMaxBytesPerline += sheet.getColumnView(cols[i]).getSize() / 256 + 1;
        int intLines = 0;
        /*
        * 一行行分析：如果一行超出最大长度，则相应的增加行高
        * */
        BufferedReader reader = new BufferedReader(new StringReader(strText));
        String strLine;
        while (null != (strLine = reader.readLine())){
            intLines++; //本身占用一行

            byte[] datas = strLine.getBytes("GBK");
            int intMinus;
            if (0 < (intMinus = datas.length - intMaxBytesPerline)){
                intLines += intMinus / intMaxBytesPerline;
                if (0 != intMinus % intMaxBytesPerline)
                    intLines++;
            } //超长
        }
        sheet.addCell(new Label(cols[0], rows, strText, format));

        if (boolean_InsertRow)//如果是新增的行，则主动设置行高
            sheet.setRowView(rows, intLines * 256); //每行高255左右

        return intLines;
    }
}
