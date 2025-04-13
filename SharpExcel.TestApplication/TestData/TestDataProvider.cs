namespace SharpExcel.TestApplication.TestData;

public static class TestDataProvider
{
    public static List<TestExportModel> GetTestData()
    {
        return new() 
        {
            new() { Id  = 0, FirstName = "John", LastName = "Doe", Budget = 2400.34m, Email = "john.doe@example.com", TestDepartment = TestDepartment.Unknown, Status = TestStatus.Employed },
            new() { Id  = 1, FirstName = "Jane", LastName = "Doe", Budget = -200.42m, Email = "jane.doe@example.com", TestDepartment = TestDepartment.ValueB, Status = TestStatus.Fired },
            new() { Id  = 2, FirstName = "John", LastName = "Neutron", Budget = 0.0m, Email = null, TestDepartment = TestDepartment.ValueB, Status = TestStatus.Employed  },
            new() { Id  = 3, FirstName = "Ash", LastName = "Ketchum", Budget = 69m, Email = null, TestDepartment = TestDepartment.ValueC, Status = TestStatus.Fired  },
            new() { Id  = 4, FirstName = "Inspector", LastName = "Gadget", Budget = 1337m, Email = "gogogadget@example.com", TestDepartment = TestDepartment.ValueC, Status = TestStatus.Employed  },
            new() { Id  = 5, FirstName = "Mindy", LastName = "", Budget = 2400.34m, Email = "mmouse@example.com", TestDepartment = TestDepartment.ValueA, Status = TestStatus.Employed  },
            new() { Id  = 6, FirstName = "ThisIsLongerThan10", LastName = "Mouse", Budget = 2400.34m, Email = "mmouse@example.com", TestDepartment = TestDepartment.ValueA, Status = TestStatus.Employed  },
            new() { Id  = 7, FirstName = "Name", LastName = "LasName", Budget = 2400.34m, Email = null, TestDepartment = TestDepartment.ValueB, Status = TestStatus.Employed  },
        };
    }
}