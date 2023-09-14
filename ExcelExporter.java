//To use this function
//1. Map the fieldOrder according to your java bean property
//2. Import an excel template
//3. Provide path to save
//4. Provide last row value if needed
//** Only used for basic excel export

public class ExcelExporter {
    public static <T> boolean exportToExcel(List<T> objects, String filePath, Class<T> clazz, String[] fieldOrder, String template, HashMap<Integer, String> lastRowValue ) throws IOException {
    	boolean success = false;
    	FileInputStream fsIP= new FileInputStream(new File(template));		
    	Workbook workbook = new XSSFWorkbook(fsIP);
        Sheet sheet = workbook.getSheetAt(0);

        int rowNum = 1;
        for (T obj : objects) {
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;
            for (String fieldName : fieldOrder) {
                try {
                	Method method = clazz.getMethod("get" + Character.toUpperCase(fieldName.charAt(0)) + fieldName.substring(1));
                    Object value = method.invoke(obj);
                    Cell cell = row.createCell(cellNum++);
                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Integer) {
                        cell.setCellValue((Integer) value);
                    } else if (value instanceof Date) {
                    	cell.setCellValue((Date) value);
                    }
                } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
        if(lastRowValue!=null) {
        	Row lastRow = sheet.getRow(rowNum);
            for (HashMap.Entry<Integer, String> entry : lastRowValue.entrySet()) {
                int columnIndex = entry.getKey();
                Object value = entry.getValue();
                lastRow.createCell(columnIndex).setCellValue(value.toString());
            }
        }
        
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
            success = true;
        } catch (IOException e) {
            e.printStackTrace();
        }
		return success;
    }
}