using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Configuration;
using static exportDataToXlsx.ConsoleWriter;
using System.IO;
using System.Data;

namespace exportDataToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            string dbName = ConfigurationManager.AppSettings["DbName"];
            string selectCommand = ConfigurationManager.AppSettings["selectCommand"];
            string dataSeparator = ConfigurationManager.AppSettings["dataSeparator"];

            if (dbName==null)
            {
                Write("Please enter DbName parameter in application settings!");
                return;
            }
            
            if (selectCommand==null)
            {
                Write("Please enter selectCommand parameter in application settings!");
                return;
            }

            if (dataSeparator==null)
            {
                dataSeparator = "|";
            }

            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(currentDir, dbName);
            Write($"Read parameters are: dbname={dbName}, selectCommand={selectCommand}");
            Write($"Db file name = {dbPath}");

            try
            {
                using (SQLiteConnection connection = new SQLiteConnection())
                {
                    connection.ConnectionString = "Data Source = " + dbPath;

                    connection.Open();
                    Write("Connection to database opened!");
                    using (SQLiteCommand command = new SQLiteCommand(connection))
                    {
                        command.CommandText = selectCommand;
                        command.CommandType = CommandType.Text;
                        SQLiteDataReader reader = command.ExecuteReader();
                        StreamWriter writer = null;
                        try
                        {
                            writer = new StreamWriter (File.OpenWrite(Path.Combine(currentDir, "result.csv")),Encoding.UTF8);
                            Write("Result file is opened successfully!");
                            StringBuilder str = new StringBuilder();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                str.Append("\""+reader.GetName(i)+"\"");
                                if (i<=(reader.FieldCount-1))
                                str.Append(dataSeparator);
                            }
                            writer.WriteLine(str.ToString());
                            str.Clear();
                            while (reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    str.Append("\""+reader[i]+ "\"");
                                    if (i <= (reader.FieldCount - 1))
                                        str.Append(dataSeparator);
                                }
                                writer.WriteLine(str.ToString());
                                str.Clear();
                            }
                            Write("Result file is written successfully!");
                        }
                        catch(IOException ex)
                        {
                            Write("Error occured during csv file writning!");
                            Write("Exception:" + ex.ToString());
                        }
                                              
                    }
                }
            }
            catch(BadImageFormatException ex)
            {
                Write("Error occured! Db file cannot be loaded!");
                Write("Exception:"+ex.ToString());
            }
            catch (SQLiteException ex)
            {
                Write("Error in sql command!");
                Write("Exception:" + ex.ToString());

            }

            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }
        
    }
}
