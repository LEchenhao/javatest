package excel;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

 

public class ExcelOperate {

 

    public static void main(String[] args) throws Exception {

    //  清空 输出路径下(中文分词模块)已有数据，以保证本次结果纯净        中文语料
          File folderCN = new File("G:/Workspace/SentimentNew/data_trains");  // 此路径在上级调用project内
		  File[] filesCN = folderCN.listFiles();
		  for (int i=0;i<filesCN.length;i++){
		    File fileexist = filesCN[i];
		    if (fileexist.exists()){
		    	fileexist.delete();
		    }
		  }
	    

     //  清空 输出路径下已有数据，以保证本次结果纯净     英文语料时，不需要做分词，因而数据不去分词模块，直接进入LDA分析模块
	   /* 	File folder = new File("E:/WorkSpace_2014/NLPLDAYL/testdata/InputEN");
			  File[] files = folder.listFiles();
			  for (int i=0;i<files.length;i++){
			    File fileexist = files[i];
			    if (fileexist.exists()){
			    	fileexist.delete();
			    }
			  }	 */ 
		  
   //  读入 Excel 处理成  txt  ，这里可以改为批处理 
	//中文部分
	   //File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\Olym.xls");  //具体到 待抽取的文件
	   //File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\kuke.xls");  //具体到 待抽取的文件 
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\engage.xls");  //具体到 待抽取的文件
	     File file = new File("G:\\Workspace\\Post2Txt\\data\\post.xls");   // 输入的待处理的Excel 文档
	     
	 //英文部分    
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\arttalk.xls");
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\twitter_king_post.xls");
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\twitter_happiness.xls");
	    // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\WIHappiness.xls");
	   
       String[][] result = getData(file, 1);

       int rowLength = result.length;

       for(int i=0;i<rowLength;i++) {

           for(int j=0;j<result[i].length;j++) {
        	   if (j==4)             // 处理xls文档的第j列
        	  { 
	        	  System.out.print(result[i][j]+"\t\t");
	        	  //此处中文、英文语料的处理结果分开，与前呼应
	             // String filepath="E:/WorkSpace_2014/CNWordsParse/Inputdata/Auto/"+Integer.toString(i);  //中文语料时 该步骤的输出路径   作为分词的输入
	            //  String filepath="E:/WorkSpace_2014/NLPLDAYL/testdata/InputEN/"+Integer.toString(i);  //英文语料时 该步骤的输出路径   
	             
	        	  String filepath="G:/Workspace/SentimentNew/data_trains/"+Integer.toString(i);  //英文语料时 该步骤的输出路径  
	        	  FileOutputStream fout=new FileOutputStream(filepath);
	              OutputStreamWriter osw=new OutputStreamWriter(fout);
	              osw.write(result[i][j]);
	              osw.close();
        	  }
           }
           System.out.println();
       }

    }

    /**

     * 读取Excel的内容，第一维数组存储的是一行中格列的值，二维数组存储的是多少个行

     * @param file 读取数据的源Excel

     * @param ignoreRows 读取数据忽略的行数，比喻行头不需要读入 忽略的行数为1  

     * @return 读出的Excel中数据的内容

     * @throws FileNotFoundException

     * @throws IOException

     */

    public static String[][] getData(File file, int ignoreRows)

           throws FileNotFoundException, IOException {

       List<String[]> result = new ArrayList<String[]>();

       int rowSize = 0;

       BufferedInputStream in = new BufferedInputStream(new FileInputStream(

              file));

       // 打开HSSFWorkbook

       POIFSFileSystem fs = new POIFSFileSystem(in);

       HSSFWorkbook wb = new HSSFWorkbook(fs);

       HSSFCell cell = null;

       for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {

           HSSFSheet st = wb.getSheetAt(sheetIndex);

           // 第一行为标题，不取     输入文件必须有 第一行 作为表头

           for (int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum(); rowIndex++) {

              HSSFRow row = st.getRow(rowIndex);

              if (row == null) {

                  continue;

              }

              int tempRowSize = row.getLastCellNum() + 1;

              if (tempRowSize > rowSize) {

                  rowSize = tempRowSize;

              }

              String[] values = new String[rowSize];

              Arrays.fill(values, "");

              boolean hasValue = false;

              for (short columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {

                  String value = "";

                  cell = row.getCell(columnIndex);

                  if (cell != null) {

                     // 注意：一定要设成这个，否则可能会出现乱码

                     //cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                	 // cell.setEncoding(HSSFCell.ENCODING_GBK);

                     switch (cell.getCellType()) {

                     case HSSFCell.CELL_TYPE_STRING:

                         value = cell.getStringCellValue();

                         break;

                     case HSSFCell.CELL_TYPE_NUMERIC:

                         if (HSSFDateUtil.isCellDateFormatted(cell)) {

                            Date date = cell.getDateCellValue();

                            if (date != null) {

                                value = new SimpleDateFormat("yyyy-MM-dd")

                                       .format(date);

                            } else {

                                value = "";

                            }

                         } else {

                            value = new DecimalFormat("0").format(cell

                                   .getNumericCellValue());

                         }

                         break;

                     case HSSFCell.CELL_TYPE_FORMULA:

                         // 导入时如果为公式生成的数据则无值

                         if (!cell.getStringCellValue().equals("")) {

                            value = cell.getStringCellValue();

                         } else {

                            value = cell.getNumericCellValue() + "";

                         }

                         break;

                     case HSSFCell.CELL_TYPE_BLANK:

                         break;

                     case HSSFCell.CELL_TYPE_ERROR:

                         value = "";

                         break;

                     case HSSFCell.CELL_TYPE_BOOLEAN:

                         value = (cell.getBooleanCellValue() == true ? "Y"

                                : "N");

                         break;

                     default:

                         value = "";

                     }

                  }

                  if (columnIndex == 0 && value.trim().equals("")) {

                     break;

                  }

                  values[columnIndex] = rightTrim(value);

                  hasValue = true;

              }

 

              if (hasValue) {

                  result.add(values);

              }

           }

       }

       in.close();

       String[][] returnArray = new String[result.size()][rowSize];

       for (int i = 0; i < returnArray.length; i++) {

           returnArray[i] = (String[]) result.get(i);

       }

       return returnArray;

    }

   

    /**

     * 去掉字符串右边的空格

     * @param str 要处理的字符串

     * @return 处理后的字符串

     */

     public static String rightTrim(String str) {

       if (str == null) {

           return "";

       }

       int length = str.length();

       for (int i = length - 1; i >= 0; i--) {

           if (str.charAt(i) != 0x20) {

              break;

           }

           length--;

       }

       return str.substring(0, length);

    }

}
