using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using NLog;

namespace MakeDxfUpdatePartData
{
    public partial class MakeDxfExportPartDataClass
    {
        #region Огранизация логирования

        /// <summary>
        /// The message to user
        /// </summary>
        public static string MessageToUsr;
        
        static readonly Logger NLogger = LogManager.GetLogger("MakeDxfExportPartDataClass");
        

        static void LoggerDebug(string logText)
        {
            LoggerMine.WriteLine(logText);
            NLogger.Log(LogLevel.Debug, logText);
            MessageToUsr = logText;
        }


        static void LoggerError(string logText, string код, string функция)
        {
            LoggerMine.Error(logText, код, функция);
            NLogger.Log(LogLevel.Error, logText);
            MessageToUsr = logText;
        }

        static void LoggerInfo(string logText, string код, string функция)
        {
            LoggerMine.Info(logText, код, функция);
            NLogger.Log(LogLevel.Info, logText);
            MessageToUsr = logText;
        }
        
        static void LoggerFatal(string logText)
        {
            LoggerMine.WriteLine(logText);
            NLogger.Log(LogLevel.Fatal, logText);
            MessageToUsr = logText;
        }

        static void LoggerTrace(string logText)
        {
            LoggerMine.WriteLine(logText);
            NLogger.Log(LogLevel.Trace, logText);
            MessageToUsr = logText;
        }

        static void LoggerWarn(string logText)
        {
            LoggerMine.WriteLine(logText);
            NLogger.Log(LogLevel.Warn, logText);
            MessageToUsr = logText;
        }

        static void LoggerOff(string logText)
        {
            LoggerMine.WriteLine(logText);
            NLogger.Log(LogLevel.Off, logText);
            MessageToUsr = logText;
        }

        #endregion
    }

    static class LoggerMine
    {
        private const string ConnectionString = "Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin";

        private const string ClassName = "ExportPartDataClass";

        //----------------------------------------------------------
        // Статический метод записи строки в файл лога без переноса
        //----------------------------------------------------------
        public static void Write(string text)
        {
            using (var streamWriter = new StreamWriter("C:\\log.txt", true))  //Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + 
            {
                streamWriter.Write(text);
            }
        }

        //---------------------------------------------------------
        // Статический метод записи строки в файл лога с переносом
        //---------------------------------------------------------
        public static void WriteLine(string message)
        {
            using (var streamWriter = new StreamWriter("C:\\log.txt", true))
            {
                streamWriter.WriteLine("{0,-23} {1}", DateTime.Now + ":", message);
            }
        }



        public static void Error(string message, string код, string функция)
        {
            using (var streamWriter = new StreamWriter("C:\\log.txt", true))
            {
                streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Error", ClassName);
            }
            WriteToBase(message, "Error", код, ClassName, функция);
        }

        public static void Info(string message, string код, string функция)
        {
            using (var streamWriter = new StreamWriter("C:\\log.txt", true))
            {
                streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Info", ClassName);
            }
            WriteToBase(message, "Info", код, ClassName, функция);
        }

        static void WriteToBase(string описание, string тип, string код, string модуль, string функция)
        {
            using (var con = new SqlConnection(ConnectionString))
            {
                try
                {
                    con.Open();
                    var sqlCommand = new SqlCommand("AddErrorLog", con) { CommandType = CommandType.StoredProcedure };

                    var sqlParameter = sqlCommand.Parameters;

                    sqlParameter.AddWithValue("@UserName", Environment.UserName + " (" + System.Net.Dns.GetHostName() + ")");
                    sqlParameter.AddWithValue("@ErrorModule", модуль);
                    sqlParameter.AddWithValue("@ErrorMessage", описание);
                    sqlParameter.AddWithValue("@ErrorCode", код);
                    sqlParameter.AddWithValue("@ErrorTime", DateTime.Now);
                    sqlParameter.AddWithValue("@ErrorState", тип);
                    sqlParameter.AddWithValue("@ErrorFunction", функция);


                    //var returnParameter = sqlCommand.Parameters.Add("@ProjectNumber", SqlDbType.Int);
                    //returnParameter.Direction = ParameterDirection.ReturnValue;

                    sqlCommand.ExecuteNonQuery();

                    //var result = Convert.ToInt32(returnParameter.Value);

                    //switch (result)
                    //{
                    //    case 0:
                    //        MessageBox.Show("Подбор №" + Номерподбора.Text + " уже существует!");
                    //        break;
                    //}

                }
                catch (Exception exception)
                {
                    
                }
                finally
                {
                    con.Close();
                }
            }
        }
    }
}
