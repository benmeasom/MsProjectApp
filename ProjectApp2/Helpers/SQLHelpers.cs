using ProjectApp2.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace ProjectApp2.Helpers
{
    public static class SQLHelpers
    {
        static readonly string connectionString = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;

        public static void ImportTasksToDatabase(List<ExcelMap.ExcelMapTask> excelMapTasks, string tableName)
        {
            var dt = ListToDataTable(excelMapTasks);
            List<string> colsToDelete = new List<string>();

            List<string> listacolumnas = new List<string>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                //For tasks and resources first check if there is value in the table and delete if there is
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = $"SELECT COUNT(*) FROM {tableName}";
                    connection.Open();
                    int count = (int)command.ExecuteScalar();

                    if (count > 0)
                    {
                        MessageBoxResult result =
                            MessageBox.Show("There are already tasks imported to Databse, if you continue previous tasks will be deleted, Do you want to continue!",
                            "", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                        if (result == MessageBoxResult.OK)
                        {
                            using (SqlCommand command2 = connection.CreateCommand())
                            {
                                command2.CommandText = $"DELETE FROM {tableName}";
                                command2.ExecuteScalar();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Transfer Terminated!!!", "", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }
                    }
                }
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = "select c.name from sys.columns c inner join sys.tables t on " +
                                          $"t.object_id = c.object_id and t.name = '{tableName}' and t.type = 'U'";
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            listacolumnas.Add(reader.GetString(0));
                        }
                    }
                }
            }

            foreach (DataColumn column in dt.Columns)
            {
                if (!listacolumnas.Contains(column.ColumnName))
                    colsToDelete.Add(column.ColumnName);
            }

            foreach (var item in colsToDelete)
            {
                dt.Columns.Remove(item);
            }

            using (var bulkCopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                foreach (DataColumn col in dt.Columns)
                {
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                }

                bulkCopy.BulkCopyTimeout = 600;
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.WriteToServer(dt);
            }
        }

        public static DataTable ListToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        public static DataTable GetDTFromSqlString(DatabaseItems databaseItems, string sqlString)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection conn = new SqlConnection())
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectionString)
                {
                    DataSource = databaseItems.ServerName,
                    InitialCatalog = databaseItems.DatabaseName
                };
                conn.ConnectionString = builder.ConnectionString;
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sqlString;
                    conn.Open();
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dataTable);
                    }
                }
            }
            return dataTable;
        }

        public static DatabaseItems GetDefaultConStringItems()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectionString);

            DatabaseItems databaseItems = new DatabaseItems()
            {
                DatabaseName = builder.InitialCatalog,
                ServerName = builder.DataSource
            };

            return databaseItems;
        }
    }
}
