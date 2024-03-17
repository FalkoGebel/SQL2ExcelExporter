using ExporterLogicLibrary;
using ExporterLogicLibrary.Models;
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
            List<ColumnModel> columns = SqlLogic.GetColumnsForTable(_serverFromFile, "master", "spt_monitor");
            columns.Count.Should().Be(11);
        }

        [TestMethod]
        public void CallGetColumnsForSptMonitorTableInMasterDBAndGetColumns()
        {
            List<ColumnModel> expected = [
                new ColumnModel() {Name = "lastrun", Type = "datetime"},
                new ColumnModel() {Name = "cpu_busy", Type = "int"},
                new ColumnModel() {Name = "io_busy", Type = "int"},
                new ColumnModel() {Name = "idle", Type = "int"},
                new ColumnModel() {Name = "pack_received", Type = "int"},
                new ColumnModel() {Name = "pack_sent", Type = "int"},
                new ColumnModel() {Name = "connections", Type = "int"},
                new ColumnModel() {Name = "pack_errors", Type = "int"},
                new ColumnModel() {Name = "total_read", Type = "int"},
                new ColumnModel() {Name = "total_write", Type = "int"},
                new ColumnModel() {Name = "total_errors", Type = "int"}
            ];

            List<ColumnModel> columns = SqlLogic.GetColumnsForTable(_serverFromFile, "master", "spt_monitor");
            columns.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        public void CallGetContentForMSreplication_optionsTableInMasterDBAndGetCorrectNumberOfEntries()
        {
            List<List<string>> lines = SqlLogic.GetContentForTable(_serverFromFile, "master", "MSreplication_options");
            lines.Count.Should().Be(3);
        }

        [TestMethod]
        public void CallGetContentForCompanyTableInDemoDatabaseNAV_11_0_DBAndGetCorrectNumberOfEntries()
        {
            List<List<string>> lines = SqlLogic.GetContentForTable(_serverFromFile, "Demo Database NAV (11-0)", "Company");
            lines.Count.Should().Be(1);
        }

        [TestMethod]
        public void CallGetContentForMSreplication_optionsTableInMasterDBAndGetCorrectValues()
        {
            List<List<string>> expected = [
                ["transactional", "True", "90", "0", "0", "0"],
                ["merge", "True", "90", "0", "0", "0"],
                ["security_model", "True", "90", "0", "0", "0"]
            ];

            List<List<string>> lines = SqlLogic.GetContentForTable(_serverFromFile, "master", "MSreplication_options");
            lines.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        public void CallGetContentForMSreplication_optionsTableInMasterDBWithSubsetOfColumnsAndGetCorrectValues()
        {
            List<string> columns = ["optname", "value", "major_version", "install_failures"];

            List<List<string>> expected = [
                ["transactional", "True", "90", "0"],
                ["merge", "True", "90", "0"],
                ["security_model", "True", "90", "0"]
            ];

            List<List<string>> lines = SqlLogic.GetContentForTable(_serverFromFile, "master", "MSreplication_options",
                columns);
            lines.Should().BeEquivalentTo(expected);
        }

        [TestMethod]
        public void CallGetContentForMSreplication_optionsTableInMasterDBWithColumsListAndGetCorrectValues()
        {
            List<List<string>> expected = [
                ["transactional", "True", "90", "0", "0", "0"],
                ["merge", "True", "90", "0", "0", "0"],
                ["security_model", "True", "90", "0", "0", "0"]
            ];

            List<ColumnModel> columns = SqlLogic.GetColumnsForTable(_serverFromFile, "master", "MSreplication_options");

            List<List<string>> lines = SqlLogic.GetContentForTable(_serverFromFile, "master", "MSreplication_options",
                columns.Select(c => c.Name).ToList());
            lines.Should().BeEquivalentTo(expected);
        }
    }
}
