using ConsoleApp8.Model;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using ConsoleApp8.Utils;


namespace ConsoleApp8
{
    public class Excel : AppSettings
    {
        public Excel()
        {
            GetExcelFile();
        }
        public string GetExcelFile()
        {
            try
            {
                string connectionString = @"Data Source=DESKTOP-SLB3J82;Initial Catalog=employeejob;Integrated Security=true;TrustServerCertificate=true";
                string filePath = @"E:\documents-importants\excelsheet.xlsx";
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                Console.WriteLine("Excel Data Reading ...");

                List<Excelsheets> validData = new List<Excelsheets>();
                List<Excelsheets> invalidData = new List<Excelsheets>();

                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    int rowCount = worksheet.Dimension.Rows;
                    int columnCount = worksheet.Dimension.Columns;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        Excelsheets excelSheets = new Excelsheets();
                        var empId = worksheet.Cells[row, 1].Value;
                        var firstName = worksheet.Cells[row, 2].Value?.ToString();
                        var lastName = worksheet.Cells[row, 3].Value?.ToString();
                        var salary = worksheet.Cells[row, 4].Value?.ToString();
                        var country = worksheet.Cells[row, 5].Value?.ToString();
                        var age = worksheet.Cells[row, 6].Value?.ToString();
                        var dates = worksheet.Cells[row, 7].Value?.ToString();
                        var description = worksheet.Cells[row, 8].Value?.ToString();

                        bool isValid = true;
                        if (empId == null) { Console.WriteLine("Id not found", row); isValid = false; }
                        if (string.IsNullOrWhiteSpace(firstName)) { Console.WriteLine("firstName not found", row); isValid = false; }
                        if (string.IsNullOrWhiteSpace(lastName)) { Console.WriteLine("lastName not found", row); isValid = false; }
                        if (salary == null) { Console.WriteLine("salary not found", row); isValid = false; }
                        if (string.IsNullOrWhiteSpace(country)) { Console.WriteLine("country not found", row); isValid = false; }
                        if (string.IsNullOrWhiteSpace(age)) { Console.WriteLine("age not found"); isValid = false; }
                        if (string.IsNullOrWhiteSpace(dates)) { Console.WriteLine("dates not found", row); isValid = false; }
                        if (string.IsNullOrWhiteSpace(description)) { Console.WriteLine("description not found", row); isValid = false; }

                        if (!isValid)
                        {
                            Excelsheets InvalidAddress = new Excelsheets()
                            {
                                EmpId = Convert.ToInt16(empId),
                                FirstName = firstName,
                                LastName = lastName,
                                Salary = Convert.ToInt16(salary),
                                Country = country,
                                Age = age,
                                Dates = Convert.ToDateTime(dates),
                                Description = description
                            };
                            invalidData.Add(InvalidAddress);
                        }
                        else
                        {
                            excelSheets.EmpId = Convert.ToInt32(empId);
                            excelSheets.FirstName = firstName;
                            excelSheets.LastName = lastName;
                            excelSheets.Salary = Convert.ToInt16(salary);
                            excelSheets.Country = country;
                            excelSheets.Age = age;
                            excelSheets.Dates = Convert.ToDateTime(dates);
                            excelSheets.Description = description;
                            validData.Add(excelSheets);
                        }
                    }
                }
                Console.WriteLine("Valid rows");
                Console.WriteLine("--------------------------------------");
                foreach (var item in validData)
                {
                    Console.WriteLine($"{item.EmpId},{item.FirstName},{item.LastName},{item.Salary},{item.Country},{item.Age},{item.Dates},{item.Description}");
                }

                Console.WriteLine("Invaild rows");
                Console.WriteLine("--------------------------------------");
                foreach (var item in invalidData)
                {
                    Console.WriteLine($"{item.EmpId},{item.FirstName},{item.LastName},{item.Salary},{item.Country},{item.Age},{item.Dates},{item.Description}");
                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    foreach (var item in validData)
                    {
                        bool dataExits = false;
                        string selectQuery = @"SELECT COUNT(*) FROM Employeejob1 WHERE EmpId = @EmpId";
                        using (SqlCommand command = new SqlCommand(selectQuery, connection))
                        {
                            command.Parameters.AddWithValue("@EmpId", item.EmpId);
                            int countData = (int)command.ExecuteScalar();
                            if (countData > 0)
                            {
                                dataExits = true;
                            }
                        }
                        if (dataExits)
                        {
                            string updateQuery = "UPDATE Employeejob1 SET FirstName=@FirstName, LastName=@LastName, Salary=@Salary, Country=@Country, Age=@Age, Dates=@Dates," +
                                                  " Description=@Description WHERE EmpId=@EmpId";
                            using (SqlCommand command1 = new SqlCommand(updateQuery, connection))
                            {
                                command1.Parameters.AddWithValue("@EmpId", item.EmpId);
                                command1.Parameters.AddWithValue("@FirstName", item.FirstName);
                                command1.Parameters.AddWithValue("@LastName", item.LastName);
                                command1.Parameters.AddWithValue("@Salary", item.Salary);
                                command1.Parameters.AddWithValue("@Country", item.Country);
                                command1.Parameters.AddWithValue("@Age", item.Age);
                                command1.Parameters.AddWithValue("@Dates", item.Dates);
                                command1.Parameters.AddWithValue("@Description", item.Description);
                                int rowAffected = command1.ExecuteNonQuery();
                                if (rowAffected > 0)
                                {
                                    Console.WriteLine("This Row is Update in a Database..");
                                }
                                else
                                {
                                    Console.WriteLine("This Row is Not Updated");
                                }
                            }
                        }
                        else
                        {
                            string insertQuery = "INSERT INTO Employeejob1 (EmpId, FirstName, LastName, Salary, Country, Age, Dates, Description) VALUES (@EmpId, @FirstName," +
                                                    " @LastName, @Salary, @Country, @Age, @Dates, @Description)";
                            using (SqlCommand command1 = new SqlCommand(insertQuery, connection))
                            {
                                command1.Parameters.AddWithValue("@EmpId", item.EmpId);
                                command1.Parameters.AddWithValue("@FirstName", item.FirstName);
                                command1.Parameters.AddWithValue("@LastName", item.LastName);
                                command1.Parameters.AddWithValue("@Salary", item.Salary);
                                command1.Parameters.AddWithValue("@Country", item.Country);
                                command1.Parameters.AddWithValue("@Age", item.Age);
                                command1.Parameters.AddWithValue("@Dates", item.Dates);
                                command1.Parameters.AddWithValue("@Description", item.Description);
                                int rowAffected = command1.ExecuteNonQuery();
                                if (rowAffected > 0)
                                {
                                    Console.WriteLine("Insert row in a Database...");
                                }
                                else
                                {
                                    Console.WriteLine("This Row is Not inserted...");
                                }
                            }
                        }
                    }
                    connection.Close();
                }


            }
            catch (Exception Ex)
            {

                Console.WriteLine(Ex.Message);
            }

            return "Ok";
        }
    }
}



