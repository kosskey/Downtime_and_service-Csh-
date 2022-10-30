using Microsoft.VisualStudio.TestTools.UnitTesting;
using Date;

namespace UnitTest;

[TestClass]
public class UnitTest1
{
    [TestMethod]
    public void TestMethod_Date()
    {
        //arrange
        string date = "2022.10.27";
        string[] commands = { "day", "month", "year", "d_full", "d_briefly" };
        string[] results = { "27", "10", "2022", "27.10.2022", "27.10" };

        var q = new Date.Date(date);
        foreach (var command in commands)
        {
            //act
            string result = q.day;
            
            //assert
            Assert.AreEqual("25", result);
        }
    }
}