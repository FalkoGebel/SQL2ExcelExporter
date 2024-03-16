using ExporterLogicLibrary;
using FluentAssertions;
using Microsoft.Data.SqlClient;

namespace ExporterLogicTests
{
    [TestClass]
    public class SqlLogicTests
    {
        private static string _serverFromFile = "";

        [ClassInitialize]
#pragma warning disable IDE0060 // Nicht verwendete Parameter entfernen
        public static void ClassInitialize(TestContext context)
#pragma warning restore IDE0060 // Nicht verwendete Parameter entfernen
        {
            string[] connectionFileData = File.ReadAllLines(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Sql2ExcelExporterTest\\SqlLogicTestsConnectionData.txt");

            _serverFromFile = connectionFileData[0];
        }

        [TestMethod]
        public void CallGetDatabasesWithoutServerAndException()
        {
            Action act = () => SqlLogic.GetDatabasesFromServer("");

            act.Should().Throw<ArgumentException>().WithMessage("No server specified");
        }

        [TestMethod]
        public void CallGetDatabasesWithWrongServerAndException()
        {
            Action act = () => SqlLogic.GetDatabasesFromServer("WRONGSERVER");

            act.Should().Throw<SqlException>().WithMessage("*error: 40 - *");
        }

        [TestMethod]
        public void CallGetDatabasesWithCorrectServerAndGetCorrectNumberOfDatabases()
        {
            List<string> databases = SqlLogic.GetDatabasesFromServer(_serverFromFile);
            databases.Count.Should().Be(8);
        }

        [TestMethod]
        public void CallGetDatabasesWithCorrectServerAndFindStandardDatabases()
        {
            List<string> databases = SqlLogic.GetDatabasesFromServer(_serverFromFile);
            databases.Should().Contain("master");
            databases.Should().Contain("tempdb");
            databases.Should().Contain("model");
            databases.Should().Contain("msdb");
        }
    }
}
