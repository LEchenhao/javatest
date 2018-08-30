




       // 娉ㄦ剰锛氫竴瀹氳璁炬垚杩欎釜锛屽惁鍒欏彲鑳戒細鍑虹幇涔辩爜

       //cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	 // cell.setEncoding(HSSFCell.ENCODING_GBK);

       switch (cell.getCellType()) {

       case HSSFCell.CELL_TYPE_STRING:

           value = cell.getStringCellValue();

           break;
           
       case HSSFCell.CELL_TYPE_FORMULA:

           // 瀵煎叆鏃跺鏋滀负鍏紡鐢熸垚鐨勬暟鎹垯鏃犲�

           if (!cell.getStringCellValue().equals("")) {

              value = cell.getStringCellValue();

           } else {
>>>>>>> branch 'master' of https://github.com/LEchenhao/javatest.git
faafaf
