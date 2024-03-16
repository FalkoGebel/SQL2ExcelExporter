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

        [TestMethod]
        public void CallGetTablesWithoutServerAndException()
        {
            Action act = () => SqlLogic.GetTablesForDatabase("", "");

            act.Should().Throw<ArgumentException>().WithMessage("No server specified");
        }

        [TestMethod]
        public void CallGetTablesWithoutDatabaseAndException()
        {
            Action act = () => SqlLogic.GetTablesForDatabase(_serverFromFile, "");

            act.Should().Throw<ArgumentException>().WithMessage("No database specified");
        }

        [TestMethod]
        public void CallGetTablesWithInvalidDatabaseAndException()
        {
            Action act = () => SqlLogic.GetTablesForDatabase(_serverFromFile, "INVALID_DB");

            act.Should().Throw<SqlException>().WithMessage("*INVALID_DB*");
        }

        [TestMethod]
        public void CallGetTablesForMasterDBAndGetCorrectNumberOfTables()
        {
            List<string> tables = SqlLogic.GetTablesForDatabase(_serverFromFile, "master");
            tables.Count.Should().Be(5);
        }

        [TestMethod]
        public void CallGetTablesForMasterDBAndGetCorrectTables()
        {
            List<string> expected = ["spt_fallback_db", "spt_fallback_dev", "spt_fallback_usg", "spt_monitor", "MSreplication_options"];
            List<string> tables = SqlLogic.GetTablesForDatabase(_serverFromFile, "master");
            tables.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        public void CallGetColumnsWithoutServerAndException()
        {
            Action act = () => SqlLogic.GetColumnsForTable("", "", "t");

            act.Should().Throw<ArgumentException>().WithMessage("No server specified");
        }

        [TestMethod]
        public void CallGetColumnsWithoutDatabaseAndException()
        {
            Action act = () => SqlLogic.GetColumnsForTable(_serverFromFile, "", "t");

            act.Should().Throw<ArgumentException>().WithMessage("No database specified");
        }

        [TestMethod]
        public void CallGetColumnsWithoutTableAndException()
        {
            Action act = () => SqlLogic.GetColumnsForTable(_serverFromFile, "master", "");

            act.Should().Throw<ArgumentException>().WithMessage("No table specified");
        }

        [TestMethod]
        public void CallGetColumnsForSptMonitorTableInMasterDBAndGetCorrectNumberOfColumns()
        {
            List<string> columns = SqlLogic.GetColumnsForTable(_serverFromFile, "master", "spt_monitor");
            columns.Count.Should().Be(11);
        }

        [TestMethod]
        public void CallGetColumnsForSptMonitorTableInMasterDBAndGetColumns()
        {
            List<string> expected = ["lastrun", "cpu_busy", "io_busy", "idle", "pack_received", "pack_sent", "connections",
                "pack_errors", "total_read", "total_write", "total_errors"];
            List<string> columns = SqlLogic.GetColumnsForTable(_serverFromFile, "master", "spt_monitor");
            columns.Should().BeEquivalentTo(expected);
        }
    }
}
