package exl;

import com.exlim.Exl;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.util.List;

public class UserModelTest {

    @Test
    public void userTest(){
        Exl exl = new Exl();
        String path = "src/test/resources/TestData.xlsx";
        List<User> users = exl.read(User.class, path);
        for (User user : users) {
            System.out.println(user.toString());
        }
        Assert.assertTrue(users.size()>0);
    }
}
