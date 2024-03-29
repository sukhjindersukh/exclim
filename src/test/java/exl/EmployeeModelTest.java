package exl;

import com.exlim.Exl;
import com.exlim.Recordset;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.util.List;

public class EmployeeModelTest {

    @Test
    public void employeeTest() throws Exception {
        Exl exl = new Exl();
        //If you want to change date format
        exl.setDateDataFormat("MM-dd-yyy");
        String path = "src/test/resources/TestData.xlsx";
        List<Employee> employees = exl.read(Employee.class, path);
        for (Employee employee : employees) {
            System.out.println(employee.toString());
        }
        Assert.assertTrue(employees.size()>0);

        //Other way
        //Exl exl = new Exl();
        exl.openWorkbook(path);
        Recordset recordset =exl.getRecords("Employee");
        exl.closeWorkbook();
        List<Recordset.Record> records = recordset.getRecords();
        for(Recordset.Record record:records){
            System.out.println(record.getValue("Name"));
        }


    }
}
