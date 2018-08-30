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

    //  濞撳懐鈹�鏉堟挸鍤捄顖氱窞娑擄拷娑擃厽鏋冮崚鍡氱槤濡�娼�瀹稿弶婀侀弫鐗堝祦閿涘奔浜掓穱婵婄槈閺堫剚顐肩紒鎾寸亯缁绢垰鍣�       娑擃厽鏋冪拠顓熸灐  浠�箞缂栫爜锛�缂栫爜
          File folderCN = new File("G:/Workspace/SentimentNew/data_trains");  // 濮濄倛鐭惧鍕躬娑撳﹦楠囩拫鍐暏project閸愶拷
		  File[] filesCN = folderCN.listFiles();
		  for (int i=0;i<filesCN.length;i++){
		    File fileexist = filesCN[i];
		    if (fileexist.exists()){
		    	fileexist.delete();
		    }
		  }
	    

     //  濞撳懐鈹�鏉堟挸鍤捄顖氱窞娑撳鍑￠張澶嬫殶閹诡噯绱濇禒銉ょ箽鐠囦焦婀板▎锛勭波閺嬫粎鍑介崙锟�   閼昏鲸鏋冪拠顓熸灐閺冭绱濇稉宥夋付鐟曚礁浠涢崚鍡氱槤閿涘苯娲滈懓灞炬殶閹诡喕绗夐崢璇插瀻鐠囧秵膩閸ф绱濋惄瀛樺复鏉╂稑鍙哃DA閸掑棙鐎藉Ο鈥虫健
	   /* 	File folder = new File("E:/WorkSpace_2014/NLPLDAYL/testdata/InputEN");
			  File[] files = folder.listFiles();
			  for (int i=0;i<files.length;i++){
			    File fileexist = files[i];
			    if (fileexist.exists()){
			    	fileexist.delete();
			    }
			  }	 */ 
		  
   //  鐠囪鍙�Excel 婢跺嫮鎮婇幋锟�txt  閿涘矁绻栭柌灞藉讲娴犮儲鏁兼稉鐑樺婢跺嫮鎮�
	//娑擃厽鏋冮柈銊ュ瀻
	   //File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\Olym.xls");  //閸忚渹缍嬮崚锟藉鍛▕閸欐牜娈戦弬鍥︽
	   //File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\kuke.xls");  //閸忚渹缍嬮崚锟藉鍛▕閸欐牜娈戦弬鍥︽ 
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\engage.xls");  //閸忚渹缍嬮崚锟藉鍛▕閸欐牜娈戦弬鍥︽
	     File file = new File("G:\\Workspace\\Post2Txt\\data\\post.xls");   // 鏉堟挸鍙嗛惃鍕窡婢跺嫮鎮婇惃鍑焫cel 閺傚洦銆�
	     
	 //閼昏鲸鏋冮柈銊ュ瀻    
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\arttalk.xls");
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\twitter_king_post.xls");
	  // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\twitter_happiness.xls");
	    // File file = new File("E:\\WorkSpace_2014\\Post2Txt\\data\\WIHappiness.xls");
	   
       String[][] result = getData(file, 1);

       int rowLength = result.length;

       for(int i=0;i<rowLength;i++) {

           for(int j=0;j<result[i].length;j++) {
        	   if (j==4)             // 婢跺嫮鎮妜ls閺傚洦銆傞惃鍕儑j閸掞拷
        	  { 
	        	  System.out.print(result[i][j]+"\t\t");
	        	  //濮濄倕顦╂稉顓熸瀮閵嗕浇瀚抽弬鍥嚔閺傛瑧娈戞径鍕倞缂佹挻鐏夐崚鍡楃磻閿涘奔绗岄崜宥呮嚑鎼达拷
	             // String filepath="E:/WorkSpace_2014/CNWordsParse/Inputdata/Auto/"+Integer.toString(i);  //娑擃厽鏋冪拠顓熸灐閺冿拷鐠囥儲顒炴銈囨畱鏉堟挸鍤捄顖氱窞   娴ｆ粈璐熼崚鍡氱槤閻ㄥ嫯绶崗锟�	            //  String filepath="E:/WorkSpace_2014/NLPLDAYL/testdata/InputEN/"+Integer.toString(i);  //閼昏鲸鏋冪拠顓熸灐閺冿拷鐠囥儲顒炴銈囨畱鏉堟挸鍤捄顖氱窞   
	             
	        	  String filepath="G:/Workspace/SentimentNew/data_trains/"+Integer.toString(i);  //閼昏鲸鏋冪拠顓熸灐閺冿拷鐠囥儲顒炴銈囨畱鏉堟挸鍤捄顖氱窞  
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

     * 鐠囪褰嘐xcel閻ㄥ嫬鍞寸�鐧哥礉缁楊兛绔寸紒瀛樻殶缂佸嫬鐡ㄩ崒銊ф畱閺勵垯绔寸悰灞艰厬閺嶇厧鍨惃鍕拷閿涘奔绨╃紒瀛樻殶缂佸嫬鐡ㄩ崒銊ф畱閺勵垰顦跨亸鎴滈嚋鐞涳拷

     * @param file 鐠囪褰囬弫鐗堝祦閻ㄥ嫭绨瓻xcel

     * @param ignoreRows 鐠囪褰囬弫鐗堝祦韫囩晫鏆愰惃鍕攽閺佸府绱濆В鏂挎煟鐞涘苯銇旀稉宥夋付鐟曚浇顕伴崗锟借箛鐣屾殣閻ㄥ嫯顢戦弫棰佽礋1  

     * @return 鐠囪鍤惃鍑焫cel娑擃厽鏆熼幑顔炬畱閸愬懎顔�

     * @throws FileNotFoundException

     * @throws IOException

     */

    public static String[][] getData(File file, int ignoreRows)

           throws FileNotFoundException, IOException {

       List<String[]> result = new ArrayList<String[]>();

       int rowSize = 0;

       BufferedInputStream in = new BufferedInputStream(new FileInputStream(

              file));

       // 閹垫挸绱慔SSFWorkbook

       POIFSFileSystem fs = new POIFSFileSystem(in);

       HSSFWorkbook wb = new HSSFWorkbook(fs);

       HSSFCell cell = null;

       for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {

           HSSFSheet st = wb.getSheetAt(sheetIndex);

           // 缁楊兛绔寸悰灞艰礋閺嶅洭顣介敍灞肩瑝閸欙拷    鏉堟挸鍙嗛弬鍥︽韫囧懘銆忛張锟界粭顑跨鐞涳拷娴ｆ粈璐熺悰銊ャ仈

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

                     // 濞夈劍鍓伴敍姘鐎规俺顩︾拋鐐灇鏉╂瑤閲滈敍灞芥儊閸掓瑥褰查懗鎴掔窗閸戣櫣骞囨稊杈╃垳

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

                         // 鐎电厧鍙嗛弮璺侯洤閺嬫粈璐熼崗顒�础閻㈢喐鍨氶惃鍕殶閹诡喖鍨弮鐘诧拷

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

     * 閸樼粯甯��妤冾儊娑撴彃褰告潏鍦畱缁岀儤鐗�

     * @param str 鐟曚礁顦╅悶鍡欐畱鐎涙顑佹稉锟�
     * @return 婢跺嫮鎮婇崥搴ｆ畱鐎涙顑佹稉锟�
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
