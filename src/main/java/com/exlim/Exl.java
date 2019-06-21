package com.exlim;

import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class Exl {
    private final Logger logger = Logger.getLogger(Exl.class.getSimpleName());
    private Workbook workbook;
    private Sheet sheet;
    private String dateDataFormat;

    /**
     * <h1>Set date time form of your choice.</h1>
     * <p>Default: dd-MM-yyy</p>
     *
     * @param dateDataFormat Pass date time format line dd-MM-yyyy
     */
    public void setDateDataFormat(String dateDataFormat) {
        this.dateDataFormat = dateDataFormat;
    }

    /**
     * <h1>Open workbook for further interactions</h1>
     *
     * @param filePath Pass full file path
     */
    public void openWorkbook(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            workbook = WorkbookFactory.create(fileInputStream);
        } catch (Throwable ex) {
            logger.log(Level.SEVERE, ex.getMessage());
            ex.printStackTrace();
        }
    }

    /**
     * <h1>Close current opened workbook</h1>
     */
    public void closeWorkbook(){
        try {
            if (this.workbook != null) {
                this.workbook.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * <h1>Open excel sheet </h1>
     * <p>This method will open and return excel sheet</p>
     *
     * @param strSheet Pass sheet name
     * @return
     * @throws Exception If sheet not found it will throw error
     */
    Sheet getSheet(String strSheet) throws Exception {
        if (workbook != null) {
            this.sheet = workbook.getSheet(strSheet);
            if (this.sheet != null) {
                logger.info("Given sheet available");
                return this.sheet;
            } else {
                StringBuilder builder = new StringBuilder();
                workbook.sheetIterator().forEachRemaining(sh ->
                        builder.append(sh.getSheetName() + ", ")
                );
                builder.replace(builder.length() - 2, builder.length(), "");
                logger.info("Available sheet(s): " + builder.toString());
                logger.log(Level.SEVERE, String.format("Sheet '%s' not found. Please check name of sheet", strSheet));
                throw new Exception("Given sheet not found: " + strSheet);
            }
        } else {
            throw new Exception("No workbook opened yet");
        }
    }

    /**
     * <h1>Get all rows from current opened sheet</h1>
     * <p>This method will return all the available rows in current sheet.</p>
     *
     * @return ArrayList It will return List of Row
     */
    public List<Row> getRowsFromSheet() {
        List<Row> rows = new ArrayList<>();
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        for (int i = firstRow; i <= lastRow; i++) {
            rows.add(sheet.getRow(i));
        }
        return rows;
    }

    /**
     * <h1>Get all the values as List<String> from given Row object</h1>
     *
     * @param row Pass Row to get cells values as List of String
     * @return List It will return List of String
     */
    public List<String> getCellsValues(Row row) {
        List<String> cellValues = new ArrayList<>();
        row.cellIterator().forEachRemaining(cell -> {
            switch (cell.getCellType()) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        if (this.dateDataFormat == null) {
                            this.dateDataFormat = "dd-MM-yyyy";
                        }
                        Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(this.dateDataFormat.replaceAll("mm","MM"));
                        LocalDateTime ldt = LocalDateTime.ofInstant(date.toInstant(),
                                ZoneId.systemDefault());
                        cellValues.add(ldt.format(formatter));
                    } else {
                        DecimalFormat decimalFormat = new DecimalFormat();
                        cellValues.add(decimalFormat.format(cell.getNumericCellValue()).toString().replaceAll(",", ""));
                    }
                    break;
                case BLANK:
                case _NONE:
                    cellValues.add("");
                    break;
                case BOOLEAN:
                    cellValues.add(Boolean.toString(cell.getBooleanCellValue()));
                    break;
                default:
                    cellValues.add(cell.getStringCellValue());
            }
        });
        return cellValues;
    }

    /**
     * <h>Get all the records from given sheet name</h>
     * <p>This method will return Recordset object that contains all the rows as records.
     * <p><b>Example</b></p>
     * <pre>
     * {@code
     *         Exl exl = new Exl();
     *         exl.openWorkbook(path);
     *         Recordset recordset =exl.getRecords("Employee");
     *         exl.closeWorkbook();
     *         List<Recordset.Record> records = recordset.getRecords();
     *         for(Recordset.Record record:records){
     *             System.out.println(record.getValue("Name"));
     *         }
     * }
     * </pre>
     *
     * <p>Note: This method will not close workbook automatically. Please call closeWorkbook() method</p>
     *
     * @param strSheet Pass sheet name to get Records
     * @return  Recordset Return Recordset object that contains all records and can be accessed using loop on all records and then getValue using column name
     * @throws Exception
     */
    public Recordset getRecords(String strSheet) throws Exception {
        getSheet(strSheet);
        Recordset recordset = new Recordset();
        List<Row> rows = getRowsFromSheet();

        //Set Headers from First row into recordset header map
        for (int i = 0; i < getCellsValues(rows.get(0)).size(); i++) {
            recordset.setHeader(i, getCellsValues(rows.get(0)).get(i).toUpperCase().trim());
        }

        int first = 0;
        for (Row row : rows) {
            if (first < 1) {
                //I am at first row I will go back and start from next row
                first++;
                continue;
            }

            Recordset.Record record = new Recordset.Record();
            for (int i = 0; i < getCellsValues(row).size(); i++) {
                String columnName = recordset.getHeader(i);
                String columnValue = getCellsValues(row).get(i);
                record.setKeyValue(columnName, columnValue);
            }

            //Now add current record object to Recordset object
            recordset.setRecord(record);
        }
        logger.info("Total number of rows including header # " + rows.size());
        return recordset;
    }

    /**
     * <h1>Read excel sheet with exactly same name as java pojo class</h1>
     * <p><b>Example</b></p>
     * <pre>
     *  {@code
     *         Exl exl = new Exl();
     *         String path = "src/test/resources/Data.xlsx";
     *         List<Employee> employees = exl.read(Employee.class, path);
     *         for (Employee employee : employees) {
     *             System.out.println(employee.toString());
     *         }
     *  }
     *       </pre>
     *
     * @param tClass Pass your java pojo class like in example Employee
     * @param filePath
     * @param
     * @return It will return List of your java pojo in our example Employee
     */
    public <T> List<T> read(Class<T> tClass, String filePath) {
        final List<T> recordsAsClass = new ArrayList<T>();

        try {
            String className = tClass.getSimpleName();
            openWorkbook(filePath);
            Recordset records = getRecords(className);

            //Create a Object of Above class and add to > recordsAsClass
            for (Recordset.Record record : records.getRecords()) {
                T type = tClass.newInstance();
                Map<String, Field> fieldMap = getClassVariableNames(type.getClass(), records);

                for (String s : fieldMap.keySet()) {
                    //Set value for current field
                    Field currentField = fieldMap.get(s);
                    currentField.setAccessible(true);
                    if (currentField.isAccessible()) {
                        currentField.set(type, record.getValue(s));
                    }
                }
                recordsAsClass.add(type);
            }

        } catch (Throwable e) {
            e.printStackTrace();
        } finally {
            closeWorkbook();
        }

        return recordsAsClass;
    }

    private <T> Map<String, Field> getClassVariableNames(Class<T> tClass, Recordset records) throws NoSuchFieldException {
        Set<String> variableNames = new HashSet<>();
        //System.out.println("------------Column names are-------------");
        for (Field field : tClass.getDeclaredFields()) {
            variableNames.add(field.getName());
        }

        Set<String> sheetHeader = new HashSet<>();
        for (String value : records.getHeadersMap().values()) {
            sheetHeader.add(value);
        }
        Map<String, Field> fieldMap = new HashMap<>();
        for (String classVariableName : variableNames) {
            if (sheetHeader.contains(classVariableName.toUpperCase())) {
                fieldMap.put(classVariableName.toUpperCase(), tClass.getDeclaredField(classVariableName));
            }
        }
        return fieldMap;
    }
}

