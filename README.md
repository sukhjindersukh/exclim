# Welcome to exclim!
This is excel slim library to read excel sheet as Java POJO(Plain Old Java Object).

# Add dependency for your project
```xml
<dependency>
  <groupId>com.github.sukhjindersukh</groupId>
  <artifactId>Exlim</artifactId>
  <version>1.1</version>
</dependency>
```

# How to use it
1. Create a Excel sheet let say **Employee** with fields   
*Name* and *DOB*.

| Name               |DOB
|----------------|-------------------------------
|Employee_1 |25-05-1989
|Employee_2 |15-02-1980
|Employee_3 |02-05-2000

## Now read all data in java
2. Create a simple java class with exactly same name of your sheet in our case **Employee**

  ```java
      public class Employee {  
          String Name, DOB;  
      
        public String getName() {  
            return Name;  
        }  
      
        public void setName(String name) {  
            Name = name;  
        }  
      
        public String getDOB() {  
            return DOB;  
        }  
      
        public void setDOB(String DOB) {  
            this.DOB = DOB;  
        }  
      
        @Override  
      public String toString() {  
            return "Employee{" +  
                    "Name='" + Name + '\'' +  
                    ", DOB='" + DOB + '\'' +  
                    '}';  
        }  
    }
```
### If you dont want to use getter setter you can use [project lombok](https://projectlombok.org) dependency to reduce your getter setter and toString code.
```xml
       <dependency>
            <groupId>org.projectlombok</groupId>
            <artifactId>lombok</artifactId>
            <version>1.18.8</version>
            <scope>provided</scope>
        </dependency>
```
### Aftre addition of library write your java POJO(Plain Old Java Object) just use annotation @Data thats it. You are ready to use 
 ```java
      @Data
      public class Employee {  
          String Name, DOB;  
     }
```



3. Create another java class with  to test our **Employee**  class is working or not


```java
public class EmployeeModelTest {
@Test  
public void employeeTest(){  
    Exl exl = new Exl();  
    exl.setDateDataFormat("MM-dd-yyy");  
    String path = "src/test/resources/TestData.xlsx";  
    List<Employee> employees = exl.read(Employee.class, path);  
    for (Employee employee : employees) {  
        System.out.println(employee.toString());  
  
    }  
    Assert.assertTrue(employees.size()>0);  
}

}
```

> **Default date time format is dd-MM-yyyy** 

# One another way to use it
```java
              Exl exl = new Exl();
              exl.openWorkbook(path);
              Recordset recordset =exl.getRecords("Employee");
              exl.closeWorkbook();
              List<Recordset.Record> records = recordset.getRecords();
              for(Recordset.Record record:records){
                  System.out.println(record.getValue("Name"));
              }
```

# Generate your jar!
>mvn clean package

# Thanks for coming here! Feedback highly appreciated.
# If you like it then hit :sun_with_face:
