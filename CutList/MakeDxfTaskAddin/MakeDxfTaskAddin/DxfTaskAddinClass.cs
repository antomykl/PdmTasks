using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using EdmLib;
using MakeDxfUpdatePartData;

namespace ExportDataAddin
{
    /// <summary>
    /// Класс для выгрузки разверток, данных деталей в базу, сохранения чертежей в PDF...
    /// </summary>
    public class ExportSwData : IEdmAddIn5  //http://help.solidworks.com/2014/English/api/epdmapi/Schedule_Task_to_Find_Files_in_State_Addin_Example_CSharp.htm
    {
        private int _currentVer;

        #region Logger

        /// <summary>
        /// The message to usr
        /// </summary>
        public static string MessageToUsr;

        static void LoggerInfo(string logText, string код, string функция)
        {
            LoggerMine.Info(logText, код, функция);
            MessageToUsr = logText;
        }

        static void LoggerError(string logText, string код, string функция)
        {
            LoggerMine.Error(logText, код, функция);
            MessageToUsr = logText;
        }

        #endregion

        /// <summary>
        /// Gets the add in information.
        /// </summary>
        /// <param name="poInfo">The po information.</param>
        /// <param name="poVault">The po vault.</param>
        /// <param name="poCmdMgr">The po command MGR.</param>
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            try
            {
                const int ver = 18;
                poInfo.mbsAddInName = "Export parts data Task Add-In ver." + ver;
                poInfo.mbsCompany = "Vents";
                poInfo.mbsDescription = "Получение разверток для деталей из листового металла по всем конфигурациям.\nВыгрузка данных детали в базу SQL и XML файлы.\nСохранение всех конфигураций детали в eDrawings part формате.";
                poInfo.mlAddInVersion = ver;
                _currentVer = poInfo.mlAddInVersion;

                //Minimum SolidWorks Enterprise PDM version
                //needed for C# Task Add-Ins is 10.0
                poInfo.mlRequiredVersionMajor = 10;
                poInfo.mlRequiredVersionMinor = 0;

                //Register this add-in as a task add-in
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
                //Register this add-in to be called when
                //selected as a task in the Administration tool
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
                
            }
            catch (COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Called when [command].
        /// </summary>
        /// <param name="poCmd">The po command.</param>
        /// <param name="ppoData">The ppo data.</param>
        public void OnCmd(ref EdmCmd poCmd, ref Array ppoData)
        {
            try
            {
                // PauseToAttachProcess(poCmd.meCmdType.ToString());
                switch (poCmd.meCmdType)
                {
                    case EdmCmdType.EdmCmd_TaskRun:
                        OnTaskRun(ref poCmd, ref ppoData);
                        break;
                    case EdmCmdType.EdmCmd_TaskSetup:
                        OnTaskSetup(ref poCmd, ref ppoData);
                        break;
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        readonly List<FileInfo> _filesToRegisration = new List<FileInfo>();

        private string _vaultName;

        void Registration()
        {
            if (_filesToRegisration.Count == 1)
            {
                Thread.Sleep(5000);
            }

            foreach (var fileInfo in _filesToRegisration)
            {
                LoggerInfo(String.Format("Начата регистрация файла {0} в хранилище {1}", Path.GetFileName(fileInfo.FullName), _vaultName), "", "Registration");
                var @class = new MakeDxfExportPartDataClass{ PdmBaseName = _vaultName };
                Thread.Sleep(500);
                if (@class.RegistrationPdm(fileInfo.FullName))
                {
                    Thread.Sleep(500);
                    LoggerInfo("Завершена регистрация файла " + Path.GetFileName(fileInfo.FullName), "", "Registration");
                    if (_filesToRegisration.Count == 1)
                    {
                        Thread.Sleep(5000);
                    }
                }
                else
                {
                    LoggerInfo(String.Format("Файл {0} не зарегестрирован в хранилище {1}! ", Path.GetFileName(fileInfo.FullName), _vaultName), "", "Registration");
                }
            }

            //if (_filesToRegisration.Count !=0)
            //{
            //    LoggerInfo("Завершена регистрация файлов в количестве " + _filesToRegisration.Count, "", "Registration");
            //}
            
        }

        void GetAllLocalFiles(string filePath)
        {
            var @class = new MakeDxfExportPartDataClass { PdmBaseName = _vaultName };
            Thread.Sleep(500);
            var extension = Path.GetExtension(filePath);
            if (extension == null) return;
            switch (extension.ToLower())
            {
                case ".sldasm":
                    @class.SearchDoc(Path.GetFileNameWithoutExtension(filePath),
                        MakeDxfExportPartDataClass.SwDocType.SwDocAssembly);
                    break;
                case ".slddrw":
                    @class.SearchDoc(Path.GetFileNameWithoutExtension(filePath),
                        MakeDxfExportPartDataClass.SwDocType.SwDocDrawing);
                    break;
            }
        }


        private void OnTaskRun(ref EdmCmd poCmd, ref Array ppoData)
        {
            var edmTaskInstance = (IEdmTaskInstance)poCmd.mpoExtra;
            edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_Running);
            LoggerInfo(String.Format("Задача '{1}' запущена для {0} элемента(ов) ", ppoData.Length, edmTaskInstance.TaskName), "", "OnTaskRun ver." + _currentVer);
            var vault = (IEdmVault7)poCmd.mpoVault;
            try
            {
                foreach (EdmCmdData edmCmdData in ppoData)
                {
                    var newFileName = "";
                    var filePath = "";

                    try
                    {
                        var fileId = edmCmdData.mlObjectID1;
                        var parentFolderId = edmCmdData.mlObjectID2;
                        var file = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, fileId);
                        var folder = (IEdmFolder7)vault.GetObject(EdmObjectType.EdmObject_Folder, parentFolderId);
                        folder = (IEdmFolder7)vault.GetFolderFromPath(folder.LocalPath);
                        file = folder.GetFile(file.Name);
                        file.GetFileCopy(poCmd.mlParentWnd, 0, folder.ID, (int)EdmGetFlag.EdmGet_Simple);
                        filePath = file.GetLocalPath(folder.ID);
                    }
                    catch (COMException exception)
                    {
                        LoggerError("Ошибка: " + exception.StackTrace + " Message " + exception.Message + " Message " + exception.Data, exception.ErrorCode.ToString("X"), "OnTaskRun");
                        edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode(), exception.StackTrace);
                    }

                    catch (Exception exception)
                    {
                        LoggerError("Ошибка: " + exception.StackTrace, GetHashCode().ToString("X"), "OnTaskRun");
                        edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode(), exception.StackTrace);
                    }

                    try
                    {
                        if (filePath == "") return;
                        _vaultName = vault.Name;
                        var extension = Path.GetExtension(filePath);
                        if (extension == null) continue;

                        switch (extension.ToLower())
                        {
                            case ".sldprt":
                                LoggerInfo("Начата обработка " + Path.GetFileName(filePath), "", "OnTaskRun");
                                var @class = new MakeDxfExportPartDataClass
                                {
                                    PdmBaseName = _vaultName 
                                };
                                bool isErrors;
                                @class.CreateFlattPatternUpdateCutlistAndEdrawing(filePath, out newFileName, out isErrors, false, false, false);
                                if (!isErrors)
                                {
                                    LoggerInfo("Закончена обработка " + Path.GetFileName(filePath), "", "OnTaskRun");
                                }
                                else
                                {
                                    LoggerError("Закончена обработка детали " + Path.GetFileName(filePath) + " с ошибками", "", "OnTaskRun");
                                    edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed);
                                }
                                
                                break;
                            #region Temp

                            case ".sldasm":
                                break;
                            case ".slddrw":
                                //try
                                //{
                                //    GetAllLocalFiles(filePath);

                                //    var @class = new MakeDxfUpdatePartDataClass
                                //    {
                                //        PdmBaseName = _vaultName
                                //    };
                                //    @class.SavePdfNew(filePath, out newFileName);
                                //}
                                //catch (Exception exception)
                                //{
                                //    MessageBox.Show(exception.Message);
                                //    // edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed);
                                //}
                                break;
                            default:
                                break;

                            #endregion
                        }
                    }

                    //catch (COMException exception)
                    //{
                    //    LoggerError("Ошибка: " + exception.StackTrace + exception.Message, exception.ErrorCode.ToString("X"), "OnTaskRun");
                    //    edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode(), exception.StackTrace);
                    //}

                    catch (Exception exception)
                    {
                        LoggerError("Ошибка: " + exception.StackTrace, GetHashCode().ToString("X"), "OnTaskRun");
                        edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode(), exception.StackTrace);
                    }

                    if (newFileName == "") continue;
                    _filesToRegisration.Add(new FileInfo(newFileName));
                    LoggerInfo("В список файлов для добавления в хранилище добавлен: " + Path.GetFileName(newFileName), "", "OnTaskRun");
                }
            }
            //catch (COMException exception)
            //{
            //    LoggerError("Ошибка: " + exception.StackTrace + exception.Message, exception.ErrorCode.ToString("X"), "OnTaskRun");
            //    edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode(), exception.StackTrace);
            //}
            catch (Exception exception)
            {
                LoggerError("Ошибка: " + exception.StackTrace, "", "OnTaskRun");
                edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode(), exception.StackTrace);
               // MessageBox.Show(exception.Message);
            }
            finally
            {
                Thread.Sleep(500);
                Registration();
                LoggerInfo("Обработка файлов завершена", "", "OnTaskRun");
                edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneOK);
            }
            LoggerInfo(String.Format("Задача '{1}' завершена для {0} элемента(ов) ", _filesToRegisration.Count, edmTaskInstance.TaskName), "", "OnTaskRun");
        }
        
        private static void OnTaskSetup(ref EdmCmd poCmd, ref Array ppoData)
        {
            try
            {
                //Get the property interface used to
                //access the framework
                var edmTaskProperties = (IEdmTaskProperties)poCmd.mpoExtra;
                LoggerInfo(String.Format("Установка задачи '{0}'. Пользователь '{1}'", edmTaskProperties.TaskName, edmTaskProperties.UserName), "", "OnTaskSetup");
                //Set the property flag that says you want a
                //menu item for the user to launch the task
                //and a flag to support scheduling
                //edmTaskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitExec + (int)EdmTaskFlag.EdmTask_SupportsScheduling; //(int)EdmTaskFlag.EdmTask_SupportsScheduling + (int)EdmTaskFlag.EdmTask_SupportsDetails;
                edmTaskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitForm + (int)EdmTaskFlag.EdmTask_SupportsDetails + (int)EdmTaskFlag.EdmTask_SupportsChangeState;


                //Set up the menu commands to launch this task
                var edmTaskMenuCmds = new EdmTaskMenuCmd[1];
                edmTaskMenuCmds[0].mbsMenuString = "Выгрузка данных деталей и чертежей";
                edmTaskMenuCmds[0].mbsStatusBarHelp = "Получение разверток для деталей из листового металла по всем конфигурациям.\nВыгрузка данных детали в базу SQL и XML файлы.\nСохранение всех конфигураций детали в eDrawings part формате.";
                edmTaskMenuCmds[0].mlCmdID = 1;
                edmTaskMenuCmds[0].mlEdmMenuFlags = (int)EdmMenuFlags.EdmMenu_Nothing;
                edmTaskProperties.SetMenuCmds(edmTaskMenuCmds);

                LoggerInfo(String.Format("Установка задачи '{0}' завершена успешно. Пользователь '{1}'", edmTaskProperties.TaskName, edmTaskProperties.UserName), "", "OnTaskSetup");
            }
            catch (COMException ex)
            {
                LoggerError("Ошибка: " + ex.StackTrace + ex.Message, ex.ErrorCode.ToString("X"), "OnTaskSetup");
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                LoggerError("Ошибка: " + ex.StackTrace, ppoData.Length.ToString("X"), "OnTaskSetup");
                MessageBox.Show("OnTaskSetup Error" + ex.Message + ppoData);
            }
        }
    }
    
    static class LoggerMine
    {
        private const string ConnectionString = "Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin";
        private const string ClassName = "ExportDataAddin";
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
                    MessageBox.Show("Введите корректные данные! " + exception.Message);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        //public static void WriteToBase(string описание, string тип, string код, string модуль, string функция)
        //{
        //    using (var con = new SqlConnection(ConnectionString))
        //    {
        //        try
        //        {
        //            con.Open();
        //            var sqlCommand = new SqlCommand("AddErrorLog", con) { CommandType = CommandType.StoredProcedure };
        //            sqlCommand.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = Environment.UserName + " (" + System.Net.Dns.GetHostName() + ")";
        //            sqlCommand.Parameters.Add("@ErrorModule", SqlDbType.NVarChar).Value = модуль;
        //            sqlCommand.Parameters.Add("@ErrorFunction", SqlDbType.NVarChar).Value = функция;
        //            sqlCommand.Parameters.Add("@ErrorCode", SqlDbType.NVarChar).Value = код;
        //            sqlCommand.Parameters.Add("@ErrorTime", SqlDbType.DateTime).Value = DateTime.Now;
        //            sqlCommand.Parameters.Add("@ErrorMessage", SqlDbType.NVarChar).Value = описание;
        //            sqlCommand.Parameters.Add("@ErrorState", SqlDbType.NVarChar).Value = тип;
        //            sqlCommand.ExecuteNonQuery();
        //        }
        //        catch (Exception)
        //        {
        //            return;
        //        }
        //        finally
        //        {
        //            con.Close();
        //        }
        //    }
        //}
    }
}
