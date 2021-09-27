
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class main {
    public static void main(String[] args) {
        HashMap<String,Integer> allOrders = new HashMap<>();
        try {
            XSSFWorkbook myExcel = new XSSFWorkbook(new FileInputStream(new File("C:\\Users\\alopukhin\\IdeaProjects\\ExcelPraser\\src\\main\\resources\\works.xlsx")));
            XSSFSheet mySheet = myExcel.getSheet("Orders");
            if (mySheet!=null) {
                int currentRow = 4;
                int rowCount = 0;
                String currentValue = mySheet.getRow(currentRow).getCell(0).toString();
                while (!currentValue.matches("Итого")) {
                    if (currentValue.matches(".*вку.*")) {
                        String orderName="";
                        Pattern pattern = Pattern.compile("вку\\d*");
                        Matcher matcher = pattern.matcher(currentValue);
                        if (matcher.find()) {
                            orderName = matcher.group();
                        }
                        currentRow+=2;
                        currentValue = mySheet.getRow(currentRow).getCell(0).toString();
                        int ordersCount=0;
                        ArrayList<Integer> currentOrder = new ArrayList<>();
                        while(!currentValue.matches(".*вку.*")&&!currentValue.matches("Итого")) {
                            ordersCount+=1;
                            currentRow+=1;
                            currentValue = mySheet.getRow(currentRow).getCell(0).toString();
                        }
                        currentOrder.add(ordersCount);
                        allOrders.put(orderName, ordersCount);
                    } else {
                        currentRow += 1;
                        currentValue = mySheet.getRow(currentRow).getCell(0).toString();
                    }
                }
            } else {
                System.err.println("Not found sheet");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
//        for (int i = 0; i < allOrders.size(); i++) {
//            System.err.println(allOrders.get(i).get(0));
//        }

        XSSFWorkbook myExcel = null;
        try {
            myExcel = new XSSFWorkbook(new FileInputStream(new File("C:\\Users\\alopukhin\\IdeaProjects\\ExcelPraser\\src\\main\\resources\\Statistics.xlsx")));
            XSSFSheet mySheet = myExcel.getSheet("TDSheet");


            int currentRow = 5;
            while (!mySheet.getRow(currentRow).getCell(1).toString().matches("Конец")) {
                for (Map.Entry item : allOrders.entrySet()) {
                    if (mySheet.getRow(currentRow).getCell(3).toString().matches(item.getKey().toString())) {
                        Row row = mySheet.getRow(currentRow);
                        Cell cell = row.createCell(13);
                        cell.setCellValue(item.getValue().toString());

                    }
                }
                currentRow+=1;
                System.err.println(mySheet.getRow(currentRow).getCell(1).toString());
            }
            myExcel.write(new FileOutputStream(new File("C:\\Users\\alopukhin\\IdeaProjects\\ExcelPraser\\src\\main\\resources\\Statistics.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
