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

    //  娓呯┖ 杈撳嚭璺緞涓�涓枃鍒嗚瘝妯″潡)宸叉湁鏁版嵁锛屼互淇濊瘉鏈缁撴灉绾噣        涓枃璇枡  什么编码？ 编码
          File folderCN = new File("G:/Workspace/SentimentNew/data_trains");  // 姝よ矾寰勫湪涓婄骇璋冪敤project鍐�
		  File[] filesCN = folderCN.listFiles();
		  for (int i=0;i<filesCN.length;i++){
		    File fileexist = filesCN[i];
		    if (fileexist.exists()){
		    	fileexist.delete();
		    }
		  }
	    

     //  娓呯┖ 杈撳嚭璺緞涓嬪凡鏈夋暟鎹紝浠ヤ繚璇佹湰娆＄粨鏋滅函鍑�    鑻辨枃璇枡鏃讹紝涓嶉渶瑕佸仛鍒嗚瘝锛屽洜鑰屾暟鎹笉鍘诲垎璇嶆ā鍧楋紝鐩存帴杩涘叆LDA鍒嗘瀽妯″潡
	   /* 	File folder = new File("E:/WorkSpace_2014/NLPLDAYL/testdata/InputEN");
			  File[] files = folder.listFiles();
			  for (int i=0;i<files.length;i++){
			    File fileexist = files[i];
			    if (fileexist.exists()){
			    	fileexist.delete();
			    }
			  }	 */ 
		  
   //  璇诲叆 Excel 澶勭悊鎴� txt  锛岃繖閲屽彲浠ユ敼涓烘壒澶勭悊 
	//涓枃閮ㄥ垎
	   //File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\Olym.xls");  //鍏蜂綋鍒�寰呮娊鍙栫殑鏂囦欢
	   //File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\kuke.xls");  //鍏蜂綋鍒�寰呮娊鍙栫殑鏂囦欢 
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\engage.xls");  //鍏蜂綋鍒�寰呮娊鍙栫殑鏂囦欢
	     File file = new File("G:\\Workspace\\Post2Txt\\data\\post.xls");   // 杈撳叆鐨勫緟澶勭悊鐨凟xcel 鏂囨。
	     
	 //鑻辨枃閮ㄥ垎    
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\arttalk.xls");
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\twitter_king_post.xls");
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\twitter_happiness.xls");
	    // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\WIHappiness.xls");
	   
       String[][] result = getData(file, 1);

       int rowLength = result.length;

       for(int i=0;i<rowLength;i++) {

           for(int j=0;j<result[i].length;j++) {
        	   if (j==4)             // 澶勭悊xls鏂囨。鐨勭j鍒�
        	  { 
	        	  System.out.print(result[i][j]+"\t\t");
	        	  //姝ゅ涓枃銆佽嫳鏂囪鏂欑殑澶勭悊缁撴灉鍒嗗紑锛屼笌鍓嶅懠搴�
	             // String filepath="E:/WorkSpace_2014/CNWordsParse/Inputdata/Auto/"+Integer.toString(i);  //涓枃璇枡鏃�璇ユ楠ょ殑杈撳嚭璺緞   浣滀负鍒嗚瘝鐨勮緭鍏�
	            //  String filepath="E:/WorkSpace_2014/NLPLDAYL/testdata/InputEN/"+Integer.toString(i);  //鑻辨枃璇枡鏃�璇ユ楠ょ殑杈撳嚭璺緞   
	             
	        	  String filepath="G:/Workspace/SentimentNew/data_trains/"+Integer.toString(i);  //鑻辨枃璇枡鏃�璇ユ楠ょ殑杈撳嚭璺緞  
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

     * 璇诲彇Excel鐨勫唴瀹癸紝绗竴缁存暟缁勫瓨鍌ㄧ殑鏄竴琛屼腑鏍煎垪鐨勫�锛屼簩缁存暟缁勫瓨鍌ㄧ殑鏄灏戜釜琛�

     * @param file 璇诲彇鏁版嵁鐨勬簮Excel

     * @param ignoreRows 璇诲彇鏁版嵁蹇界暐鐨勮鏁帮紝姣斿柣琛屽ご涓嶉渶瑕佽鍏�蹇界暐鐨勮鏁颁负1  

     * @return 璇诲嚭鐨凟xcel涓暟鎹殑鍐呭

     * @throws FileNotFoundException

     * @throws IOException

     */

    public static String[][] getData(File file, int ignoreRows)

           throws FileNotFoundException, IOException {

       List<String[]> result = new ArrayList<String[]>();

       int rowSize = 0;

       BufferedInputStream in = new BufferedInputStream(new FileInputStream(

              file));

       // 鎵撳紑HSSFWorkbook

       POIFSFileSystem fs = new POIFSFileSystem(in);

       HSSFWorkbook wb = new HSSFWorkbook(fs);

       HSSFCell cell = null;

       for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {

           HSSFSheet st = wb.getSheetAt(sheetIndex);

           // 绗竴琛屼负鏍囬锛屼笉鍙�    杈撳叆鏂囦欢蹇呴』鏈�绗竴琛�浣滀负琛ㄥご

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

                     // 娉ㄦ剰锛氫竴瀹氳璁炬垚杩欎釜锛屽惁鍒欏彲鑳戒細鍑虹幇涔辩爜

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

                         // 瀵煎叆鏃跺鏋滀负鍏紡鐢熸垚鐨勬暟鎹垯鏃犲�

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

     * 鍘绘帀瀛楃涓插彸杈圭殑绌烘牸

     * @param str 瑕佸鐞嗙殑瀛楃涓�

     * @return 澶勭悊鍚庣殑瀛楃涓�

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
