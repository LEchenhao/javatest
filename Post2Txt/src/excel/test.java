import new
hah
test

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
           
       case HSSFCell.CELL_TYPE_FORMULA:

           // 导入时如果为公式生成的数据则无值

           if (!cell.getStringCellValue().equals("")) {

              value = cell.getStringCellValue();

           } else {