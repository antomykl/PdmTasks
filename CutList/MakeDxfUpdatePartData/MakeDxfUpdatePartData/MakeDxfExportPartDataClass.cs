using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using EdmLib;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using View = SolidWorks.Interop.sldworks.View;

namespace MakeDxfUpdatePartData
{
    public partial class MakeDxfExportPartDataClass
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="MakeDxfExportPartDataClass"/> class.
        /// </summary>
        public MakeDxfExportPartDataClass()
        {
            _sPathToSaveDxf = TestingCode ? @"C:\Temp\" : @"\\srvkb\DXF\";
            _xmlPath = TestingCode ? @"C:\Temp\" : @"\\srvkb\SolidWorks Admin\XML\";

            _шаблонЧертежаРазверткиВнеХранилища = @"\\srvkb\SolidWorks Admin\Templates\flattpattern.drwdot";
           // _шаблонЧертежаРазвертки = "\\Библиотека проектирования\\Templates\\flattpattern.drwdot";
           // _папкаШаблонов = "\\Библиотека проектирования\\Templates\\";
            _папкаШаблонов = @"\\srvkb\SolidWorks Admin\Templates\";
            _connectionString = "Data Source=srvkb;Initial Catalog=SWPlusDB;Persist Security Info=True;User ID=sa;Password=PDMadmin;MultipleActiveResultSets=True";
        }

        #region ModelCode

        const bool TestingCode = false;

        private const bool ШаблоныВХранилище = true;
        private readonly string _sPathToSaveDxf;
        private readonly string _xmlPath;
        private int _currentVersion;
        private string _eDrwFileName;
        private readonly string _шаблонЧертежаРазвертки;
        private readonly string _шаблонЧертежаРазверткиВнеХранилища;
        private readonly string _папкаШаблонов;
        private readonly string _connectionString;

        /// <summary>
        /// Gets or sets the name of the PDM base. For example: "Vents-PDM"
        /// </summary>
        public string PdmBaseName { get; set; }

        /// <summary>
        /// Creates the flatt pattern update cutlist and edrawing.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="eDrwFileName">Name of the e DRW file.</param>
        /// <param name="isErrors">if set to <c>true</c> [is errors].</param>
        /// <param name="makeDxf">if set to <c>true</c> [make DXF].</param>
        /// <param name="makeEprt">if set to <c>true</c> [make eprt].</param>
        /// <param name="swVisible">if set to <c>true</c> [sw visible].</param>
        public void CreateFlattPatternUpdateCutlistAndEdrawing(string filePath, out string eDrwFileName, out bool isErrors, bool makeDxf, bool makeEprt, bool swVisible)
        {
            isErrors = false;

            eDrwFileName = "";

            #region Сбор информации по детали и сохранение разверток

            SldWorks swApp = null;
            try
            {
                LoggerInfo("Запущен метод для обработки детали по пути " + filePath, "", "CreateFlattPatternUpdateCutlistAndEdrawing");

                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);

                try
                {
                    IEdmFolder5 oFolder;
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                    _currentVersion = edmFile5.CurrentVersion;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при получении значения последней версии файла {0}", Path.GetFileName(filePath)), exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                #region To Delete

                //vault1.Login("kb81","1",PdmBaseName);
                // var edmFile5 = vault1.GetFileFromPath(vault1.RootFolderPath + _шаблонЧертежаРазвертки, out oFolder);
                // edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);

                #endregion

                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = swVisible };
                }
                if (swApp == null)
                {
                    isErrors = true;
                    LoggerInfo("isErrors = true на 93-й строке ", "", "");
                    return; 
                }
                try
                {
                    swApp.Visible = swVisible;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при попытке погашния {0}", Path.GetFileName(filePath)), exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }


                IModelDoc2 swModel;

                #region To Delete

                //try
                //{
                    //IEdmFolder5 oFolder;
                    //var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    //edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);

                    //_currentVersion = edmFile5.CurrentVersion;
                //    //swApp.SetUserPreferenceStringValue(((int)(swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates)), vault1.RootFolderPath + _папкаШаблонов);
                //    swApp.SetUserPreferenceStringValue(((int)(swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates)), _папкаШаблонов);
                //}
                //catch (Exception exception)
                //{
                //    LoggerError(String.Format("Ошибка: {0} Строка: {1}", exception.Message, exception.StackTrace), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                //}

                #endregion

                try
                {
                    swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    swModel.Extension.ViewDisplayRealView = false;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при обработке детали {2}: {0} Строка: {1}", exception.Message, exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 127-й строке ", "", "");
                    return;
                }

                try
                {
                    if (!IsSheetMetalPart((IPartDoc)swModel))
                    {
                        swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                        if (!makeDxf) return;
                        swApp.ExitApp();
                        swApp = null;
                        return;
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка2 при обработке детали {2}: {0} Строка: {1}", exception.Message, exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 145-й строке ", "", "");
                }

                Configuration activeconfiguration;
                string[] swModelConfNames;

                try
                {
                    activeconfiguration = (Configuration)swModel.GetActiveConfiguration();
                    swModelConfNames = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(),"", "");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 160-й строке ", "", "");
                    return;
                }
                
                try
                {
                    foreach (var name in from name in swModelConfNames
                                         let config = (Configuration)swModel.GetConfigurationByName(name)
                                         where config.IsDerived()
                                         select name)
                    {
                        try
                        {
                            swModel.DeleteConfiguration(name);
                        }
                        catch (Exception exception)
                        {
                            LoggerError(String.Format("Ошибка при удалении конфигурации '{2}' в модели '{3}': {0} Строка: {1}", exception.Message, exception.StackTrace, name, swModel.GetTitle()),
                                exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                        }
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при удалении конфигураций в модели '{2}': {0} Строка: {1}", exception.Message, exception.StackTrace, swModel.GetTitle()),
                                 exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    LoggerInfo("isErrors = true на 186-й строке ", "", "");
                    isErrors = true;
                }

                ModelDocExtension swModelDocExt;
                string[] swModelConfNames2;

                try
                {
                    swModelDocExt = swModel.Extension;
                    swModelConfNames2 = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 202-й строке ", "", "");
                    return;
                }
                
                // Проход по всем родительским конфигурациям

                var dataList = new List<DataToExport>();
                
                try
                {
                    foreach (var configName in from name in swModelConfNames2
                                               let config = (Configuration)swModel.GetConfigurationByName(name)
                                               where !config.IsDerived()
                                               select name)
                    {

                        swModel.ShowConfiguration2(configName);
                        swModel.EditRebuild3();

                        var confiData = new DataToExport { Config = configName };
                        
                        FileInfo template = null;

                        try
                        {
                            template = new FileInfo(_шаблонЧертежаРазверткиВнеХранилища);
                            Thread.Sleep(1000);
                        }
                        catch (Exception exception)
                        {
                            LoggerError("Проблемы с полцчением шаблона чертежа", exception.StackTrace, "CreateFlattPatternUpdateCutlistAndEdrawing");
                            template = new FileInfo(_шаблонЧертежаРазверткиВнеХранилища);
                            Thread.Sleep(1000);

                        }
                        finally
                        {
                            if (template == null)
                            {
                                template = new FileInfo(_шаблонЧертежаРазверткиВнеХранилища);
                            }
                        }

                        if (!template.Exists)
                        {
                            LoggerError("Не удалось найти шаблон чертежа по пути \n" + template.FullName + "\nПроверте подключение к " + template.Directory, "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                            isErrors = true;
                            LoggerInfo("isErrors = true на 229-й строке ", "", "");
                        }

                        if (swApp != null)
                        {
                            DrawingDoc swDraw = null;
                            if (makeDxf)
                            {
                                 swDraw = (DrawingDoc)
                                    swApp.INewDocument2(template.FullName, (int)swDwgPaperSizes_e.swDwgPaperA0size, 0.841, 0.594);
                                swDraw.CreateFlatPatternViewFromModelView3(swModel.GetPathName(), configName, 0.841 / 2, 0.594 / 2, 0, true, false);
                                ((IModelDoc2)swDraw).ForceRebuild3(true);    
                            }

                            try
                            {
                                swModel.EditRebuild3();
                                var swPart = (IPartDoc)swModel;

                                Feature swFeature = swPart.FirstFeature();
                                const string strSearch = "FlatPattern";
                                while (swFeature != null)
                                {
                                    var nameTypeFeature = swFeature.GetTypeName2();

                                    if (nameTypeFeature == strSearch)
                                    {
                                        swFeature.Select(true);
                                        swPart.EditUnsuppress();

                                        Feature swSubFeature = swFeature.GetFirstSubFeature();
                                        while (swSubFeature != null)
                                        {
                                            var nameTypeSubFeature = swSubFeature.GetTypeName2();

                                            if (nameTypeSubFeature == "UiBend")
                                            {
                                                swFeature.Select(true);
                                                swPart.EditUnsuppress();
                                                swModel.EditRebuild3();

                                                try
                                                {
                                                    swSubFeature.SetSuppression2(
                                                        (int) swFeatureSuppressionAction_e.swUnSuppressFeature,
                                                        (int) swInConfigurationOpts_e.swAllConfiguration,
                                                        swModelConfNames2);
                                                }
                                                catch (Exception)
                                                {
                                                   
                                                }
                                            }
                                            swSubFeature = swSubFeature.GetNextSubFeature();
                                        }
                                    }
                                    swFeature = swFeature.GetNextFeature();
                                }
                                swModel.EditRebuild3();
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                            
                            swModel.ForceRebuild3(false);

                            var swCustProp = swModelDocExt.CustomPropertyManager[configName];
                            string valOut;

                            string codMaterial;
                            swCustProp.Get4("Код материала", true, out valOut, out codMaterial);
                            confiData.КодМатериала = codMaterial;

                            string материал;
                            swCustProp.Get4("Материал", true, out valOut, out материал);
                            confiData.Материал = материал;

                            string обозначение;
                            swCustProp.Get4("Обозначение", true, out valOut, out обозначение);
                            confiData.Обозначение = обозначение;

                            var swCustPropForDescription = swModelDocExt.CustomPropertyManager[""];
                            string наименование;
                            swCustPropForDescription.Get4("Наименование", true, out valOut, out наименование);
                            confiData.Наименование = наименование;

                            var thikness = GetFromCutlist(swModel, "Толщина листового металла");

                            if (makeDxf)
                            {
                                var errors = 0;
                                var warnings = 0;
                                var newDxf = (IModelDoc2)swDraw;

                                newDxf.Extension.SaveAs(
                                    _sPathToSaveDxf + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + "-" + configName + "-" + thikness + ".dxf",  //codMaterial
                                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                    (int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, null, ref errors, ref warnings);

                                swApp.CloseDoc(Path.GetFileName(newDxf.GetPathName()));
                            }
                            

                            //UpdateCustomPropertyListFromCutList
                            const string длинаГраничнойРамкиName = "Длина граничной рамки";
                            const string ширинаГраничнойРамкиName = "Ширина граничной рамки";
                            const string толщинаЛистовогоМеталлаNAme = "Толщина листового металла";
                            const string сгибыName = "Сгибы";
                            const string площадьПокрытияName = "Площадь покрытия";

                            Feature swFeat2 = swModel.FirstFeature();
                            while (swFeat2 != null)
                            {
                                if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                                {
                                    BodyFolder swBodyFolder = swFeat2.GetSpecificFeature2();
                                    swFeat2.Select2(false, -1);
                                    swBodyFolder.SetAutomaticCutList(true);
                                    swBodyFolder.UpdateCutList();

                                    Feature swSubFeat = swFeat2.GetFirstSubFeature();
                                    while (swSubFeat != null)
                                    {
                                        if (swSubFeat.GetTypeName2() == "CutListFolder")
                                        {
                                            BodyFolder bodyFolder = swSubFeat.GetSpecificFeature2();
                                            swSubFeat.Select2(false, -1);
                                            bodyFolder.SetAutomaticCutList(true);
                                            bodyFolder.UpdateCutList();
                                            var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                            swCustPrpMgr.Add("Площадь поверхности", "Текст",
                                                "\"SW-SurfaceArea@@@Элемент списка вырезов1@" + Path.GetFileName(swModel.GetPathName()) + "\"");


                                            string длинаГраничнойРамки;
                                            swCustPrpMgr.Get4(длинаГраничнойРамкиName, true, out valOut,
                                                out длинаГраничнойРамки);
                                            swCustProp.Set(длинаГраничнойРамкиName, длинаГраничнойРамки);
                                            confiData.ДлинаГраничнойРамки = длинаГраничнойРамки;

                                            string ширинаГраничнойРамки;
                                            swCustPrpMgr.Get4(ширинаГраничнойРамкиName, true, out valOut,
                                                out ширинаГраничнойРамки);
                                            swCustProp.Set(ширинаГраничнойРамкиName, ширинаГраничнойРамки);
                                            confiData.ШиринаГраничнойРамки = ширинаГраничнойРамки;

                                            string толщинаЛистовогоМеталла;
                                            swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme, true, out valOut,
                                                out толщинаЛистовогоМеталла);
                                            swCustProp.Set(толщинаЛистовогоМеталлаNAme, толщинаЛистовогоМеталла);
                                            confiData.ТолщинаЛистовогоМеталла = толщинаЛистовогоМеталла;

                                            string сгибы;
                                            swCustPrpMgr.Get4(сгибыName, true, out valOut, out сгибы);
                                            swCustProp.Set(сгибыName, сгибы);
                                            confiData.Сгибы = сгибы;

                                            string площадьПоверхности;
                                            swCustPrpMgr.Get4("Площадь поверхности", true, out valOut,
                                                out площадьПоверхности);
                                            swCustProp.Set(площадьПокрытияName, площадьПоверхности);
                                            confiData.ПлощадьПокрытия = площадьПоверхности;
                                        }
                                        swSubFeat = swSubFeat.GetNextFeature();
                                    }
                                }
                                swFeat2 = swFeat2.GetNextFeature();
                            }
                        }
                        dataList.Add(confiData);
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "Строка 377", "");
                }

                try
                {
                    swModel.ShowConfiguration2(activeconfiguration.Name);

                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 392-й строке ", "", "");
                }

                try
                {
                    ExportDataToXmlSql(swModel, dataList);
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 403-й строке ", "", "");
                }
                
            #endregion
                
                #region Сохранение детали в eDrawing

                if (makeEprt)
                {
                   

                    string modelName;
                    try
                    {
                        modelName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                        _eDrwFileName = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + modelName + ".eprt";
                        eDrwFileName = _eDrwFileName;
                    }
                    catch (Exception exception)
                    {
                        LoggerError(exception.ToString(), "", "");
                        isErrors = true;
                        LoggerInfo("isErrors = true на 423-й строке ", "", "");
                        return;
                    }

                    try
                    {
                        // todo: удаление документов перед новым сохранением. Осуществить поиск по имени
                        var existingDocument = SearchDoc(modelName + ".eprt", SwDocType.SwDocNone);

                        if (existingDocument != "")
                        {
                            LoggerInfo(String.Format("Файл есть в базе {0} и будет удален.. ", modelName), "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                            //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                            DeleteFileFromPdm(existingDocument, PdmBaseName);
                        }
                        else
                        {
                            File.Delete(_eDrwFileName);
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            File.Delete(_eDrwFileName);
                        }
                        catch (Exception exception)
                        {
                            LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                            isErrors = true;
                            LoggerInfo("isErrors = true на 453-й строке ", "", "");
                        }
                    }

                    #region ToDelete
                    //if (new FileInfo(_eDrwFileName).Exists)
                    //{
                    //    LoggerInfo("Файл есть в базе " + swModel.GetTitle());
                    //    //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                    //    DeleteFileFromPdm(_eDrwFileName, PdmBaseName);
                    //}
                    #endregion

                    try
                    {
                        swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption,
                        (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                        swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        swModel.Extension.SaveAs(_eDrwFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                    }
                    catch (Exception exception)
                    {
                        LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                        isErrors = true;
                        LoggerInfo("isErrors = true на 478-й строке ", "", "");
                    }

                    //try
                    //{
                    //    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //    swApp.CloseDoc(modelName + ".sldprt");

                    //    if (makeDxf)
                    //    {
                    //        swApp.ExitApp();
                    //        swApp = null;
                    //    }
                    //    LoggerInfo("Обработка файла " + modelName + ".sldprt" + " успешно завершена", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    //    isErrors = false;
                    //}
                    //catch (Exception exception)
                    //{
                    //    LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    //    isErrors = true;
                    //    LoggerInfo("isErrors = true на 497-й строке ", "", "");
                    //}


                }
                
                #endregion


                try
                {
                    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    var namePrt = swApp.IActiveDoc2.GetTitle().ToLower().Contains(".sldprt")
                        ? swApp.IActiveDoc2.GetTitle()
                        : swApp.IActiveDoc2.GetTitle() + ".sldprt";
                    swApp.CloseDoc(namePrt);

                    if (makeDxf)
                    {
                        swApp.ExitApp();
                        swApp = null;
                    }
                    LoggerInfo(
                        "Обработка файла " + swApp.IActiveDoc2.GetTitle() + ".sldprt" + ".sldprt" + " успешно завершена",
                        "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = false;
                }
                catch (Exception exception)
                {
                    LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 497-й строке ", "", "");
                }
                finally
                {
                    //try
                    //{
                    //    swApp.CloseDoc(swApp.IActiveDoc2.GetTitle() + ".sldprt");
                    //}
                    //catch (Exception)
                    //{
                    //    swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //}
                }

                
            }
            catch (Exception exception)
            {
                LoggerError(String.Format("Общая ошибка метода: {0} Строка: {1} exception.Source - ", exception.Message, exception.StackTrace), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                if (swApp == null) return;
                if (makeDxf)
                {
                    swApp.ExitApp();
                }
                isErrors = true;
                LoggerInfo("isErrors = true на 506-й строке ", "", "");
            }
        }


        /// <summary>
        /// Creates the flatt pattern update cutlist and edrawing2.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="eDrwFileName">Name of the e DRW file.</param>
        /// <param name="isErrors">if set to <c>true</c> [is errors].</param>
        /// <param name="makeDxf">if set to <c>true</c> [make DXF].</param>
        /// <param name="makeEprt">if set to <c>true</c> [make eprt].</param>
        /// <param name="swVisible">if set to <c>true</c> [sw visible].</param>
        public void CreateFlattPatternUpdateCutlistAndEdrawing2(string filePath, out string eDrwFileName, out bool isErrors, bool makeDxf, bool makeEprt, bool swVisible)
        {
            isErrors = false;

            eDrwFileName = "";

            #region Сбор информации по детали и сохранение разверток

            SldWorks swApp = null;
            try
            {
                LoggerInfo("Запущен метод для обработки детали по пути " + filePath, "", "CreateFlattPatternUpdateCutlistAndEdrawing");

                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);

                try
                {
                    IEdmFolder5 oFolder;
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                    _currentVersion = edmFile5.CurrentVersion;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при получении значения последней версии файла {0}", Path.GetFileName(filePath)), exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = swVisible };
                }
                if (swApp == null)
                {
                    isErrors = true;
                    return;
                }
                try
                {
                    swApp.Visible = swVisible;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при попытке погашния {0}", Path.GetFileName(filePath)), exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                IModelDoc2 swModel;

                try
                {
                    swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    swModel.Extension.ViewDisplayRealView = false;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при обработке детали {2}: {0} Строка: {1}", exception.Message, exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    return;
                }

                try
                {
                    if (!IsSheetMetalPart((IPartDoc)swModel))
                    {
                        swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                        if (!makeDxf) return;
                        swApp.ExitApp();
                        swApp = null;
                        return;
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка2 при обработке детали {2}: {0} Строка: {1}", exception.Message, exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                }

                Configuration activeconfiguration;
                string[] swModelConfNames;

                try
                {
                    activeconfiguration = (Configuration)swModel.GetActiveConfiguration();
                    swModelConfNames = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    return;
                }

                if (!makeEprt)
                {
                    try
                    {
                        swModel.EditRebuild3();
                        var swPart = (PartDoc)swModel;
                        var arrNamesConfig = (string[])swModel.GetConfigurationNames();

                        Feature swFeature = swPart.FirstFeature();
                        const string strSearch = "FlatPattern";

                        while (swFeature != null)
                        {
                            var nameTypeFeature = swFeature.GetTypeName2();

                            if (nameTypeFeature == strSearch)
                            {
                                Feature swSubFeature = swFeature.GetFirstSubFeature();
                                while (swSubFeature != null)
                                {
                                    var nameTypeSubFeature = swSubFeature.GetTypeName2();

                                    if (nameTypeSubFeature == "UiBend")
                                    {
                                        swSubFeature.SetSuppression2(
                                            (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                                            (int)swInConfigurationOpts_e.swAllConfiguration,
                                            arrNamesConfig);
                                    }
                                    swSubFeature = swSubFeature.GetNextSubFeature();
                                }
                            }
                            swFeature = swFeature.GetNextFeature();
                        }
                        swModel.EditRebuild3();
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }



                try
                {
                    foreach (var name in from name in swModelConfNames
                                         let config = (Configuration)swModel.GetConfigurationByName(name)
                                         where config.IsDerived()
                                         select name)
                    {
                        try
                        {
                            swModel.DeleteConfiguration(name);
                        }
                        catch (Exception exception)
                        {
                            LoggerError(String.Format("Ошибка при удалении конфигурации '{2}' в модели '{3}': {0} Строка: {1}",
                                exception.Message, exception.StackTrace, name, swModel.GetTitle()),
                                exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                        }
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при удалении конфигураций в модели '{2}': {0} Строка: {1}",
                        exception.Message, exception.StackTrace, swModel.GetTitle()),
                        exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    LoggerInfo("isErrors = true на 186-й строке ", "", "");
                    isErrors = true;
                }

                ModelDocExtension swModelDocExt;
                string[] swModelConfNames2;

                try
                {
                    swModelDocExt = swModel.Extension;
                    swModelConfNames2 = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    return;
                }

                // Проход по всем родительским конфигурациям

                var dataList = new List<DataToExport>();

                try
                {
                    foreach (var configName in from name in swModelConfNames2
                                               let config = (Configuration)swModel.GetConfigurationByName(name)
                                               where !config.IsDerived()
                                               select name)
                    {

                        swModel.ShowConfiguration2(configName);
                        swModel.EditRebuild3();

                        var confiData = new DataToExport { Config = configName };
                        

                        if (swApp != null)
                        {
                            swModel.ForceRebuild3(false);

                            var swCustProp = swModelDocExt.CustomPropertyManager[configName];
                            string valOut;

                            string codMaterial;
                            swCustProp.Get4("Код материала", true, out valOut, out codMaterial);
                            confiData.КодМатериала = codMaterial;

                            string материал;
                            swCustProp.Get4("Материал", true, out valOut, out материал);
                            confiData.Материал = материал;

                            string обозначение;
                            swCustProp.Get4("Обозначение", true, out valOut, out обозначение);
                            confiData.Обозначение = обозначение;

                            var swCustPropForDescription = swModelDocExt.CustomPropertyManager[""];
                            string наименование;
                            swCustPropForDescription.Get4("Наименование", true, out valOut, out наименование);
                            confiData.Наименование = наименование;

                            var thikness = GetFromCutlist(swModel, "Толщина листового металла");


                            


                            //UpdateCustomPropertyListFromCutList
                            const string длинаГраничнойРамкиName = "Длина граничной рамки";
                            const string ширинаГраничнойРамкиName = "Ширина граничной рамки";
                            const string толщинаЛистовогоМеталлаNAme = "Толщина листового металла";
                            const string сгибыName = "Сгибы";
                            const string площадьПокрытияName = "Площадь покрытия";

                            Feature swFeat2 = swModel.FirstFeature();
                            while (swFeat2 != null)
                            {
                                if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                                {
                                    BodyFolder swBodyFolder = swFeat2.GetSpecificFeature2();
                                    swFeat2.Select2(false, -1);
                                    swBodyFolder.SetAutomaticCutList(true);
                                    swBodyFolder.UpdateCutList();

                                    Feature swSubFeat = swFeat2.GetFirstSubFeature();
                                    while (swSubFeat != null)
                                    {
                                        if (swSubFeat.GetTypeName2() == "CutListFolder")
                                        {
                                            BodyFolder bodyFolder = swSubFeat.GetSpecificFeature2();
                                            swSubFeat.Select2(false, -1);
                                            bodyFolder.SetAutomaticCutList(true);
                                            bodyFolder.UpdateCutList();
                                            var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                            swCustPrpMgr.Add("Площадь поверхности", "Текст",
                                                "\"SW-SurfaceArea@@@Элемент списка вырезов1@" + Path.GetFileName(swModel.GetPathName()) + "\"");


                                            string длинаГраничнойРамки;
                                            swCustPrpMgr.Get4(длинаГраничнойРамкиName, true, out valOut,
                                                out длинаГраничнойРамки);
                                            swCustProp.Set(длинаГраничнойРамкиName, длинаГраничнойРамки);
                                            confiData.ДлинаГраничнойРамки = длинаГраничнойРамки;

                                            string ширинаГраничнойРамки;
                                            swCustPrpMgr.Get4(ширинаГраничнойРамкиName, true, out valOut,
                                                out ширинаГраничнойРамки);
                                            swCustProp.Set(ширинаГраничнойРамкиName, ширинаГраничнойРамки);
                                            confiData.ШиринаГраничнойРамки = ширинаГраничнойРамки;

                                            string толщинаЛистовогоМеталла;
                                            swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme, true, out valOut,
                                                out толщинаЛистовогоМеталла);
                                            swCustProp.Set(толщинаЛистовогоМеталлаNAme, толщинаЛистовогоМеталла);
                                            confiData.ТолщинаЛистовогоМеталла = толщинаЛистовогоМеталла;

                                            string сгибы;
                                            swCustPrpMgr.Get4(сгибыName, true, out valOut, out сгибы);
                                            swCustProp.Set(сгибыName, сгибы);
                                            confiData.Сгибы = сгибы;

                                            string площадьПоверхности;
                                            swCustPrpMgr.Get4("Площадь поверхности", true, out valOut,
                                                out площадьПоверхности);
                                            swCustProp.Set(площадьПокрытияName, площадьПоверхности);
                                            confiData.ПлощадьПокрытия = площадьПоверхности;
                                        }
                                        swSubFeat = swSubFeat.GetNextFeature();
                                    }
                                }
                                swFeat2 = swFeat2.GetNextFeature();
                            }
                        }
                        dataList.Add(confiData);
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "Строка 377", "");
                }

                try
                {
                    swModel.ShowConfiguration2(activeconfiguration.Name);

                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 392-й строке ", "", "");
                }

                try
                {
                    ExportDataToXmlSql(swModel, dataList);
                }
                catch (Exception exception)
                {
                    LoggerError(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 403-й строке ", "", "");
                }

            #endregion

                #region Сохранение детали в eDrawing

                if (makeEprt)
                {


                    string modelName;
                    try
                    {
                        modelName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                        _eDrwFileName = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + modelName + ".eprt";
                        eDrwFileName = _eDrwFileName;
                    }
                    catch (Exception exception)
                    {
                        LoggerError(exception.ToString(), "", "");
                        isErrors = true;
                        LoggerInfo("isErrors = true на 423-й строке ", "", "");
                        return;
                    }

                    try
                    {
                        // todo: удаление документов перед новым сохранением. Осуществить поиск по имени
                        var existingDocument = SearchDoc(modelName + ".eprt", SwDocType.SwDocNone);

                        if (existingDocument != "")
                        {
                            LoggerInfo(String.Format("Файл есть в базе {0} и будет удален.. ", modelName), "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                            //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                            DeleteFileFromPdm(existingDocument, PdmBaseName);
                        }
                        else
                        {
                            File.Delete(_eDrwFileName);
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            File.Delete(_eDrwFileName);
                        }
                        catch (Exception exception)
                        {
                            LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                            isErrors = true;
                            LoggerInfo("isErrors = true на 453-й строке ", "", "");
                        }
                    }

                    #region ToDelete
                    //if (new FileInfo(_eDrwFileName).Exists)
                    //{
                    //    LoggerInfo("Файл есть в базе " + swModel.GetTitle());
                    //    //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                    //    DeleteFileFromPdm(_eDrwFileName, PdmBaseName);
                    //}
                    #endregion

                    try
                    {
                        swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption,
                        (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                        swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        swModel.Extension.SaveAs(_eDrwFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                    }
                    catch (Exception exception)
                    {
                        LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                        isErrors = true;
                        LoggerInfo("isErrors = true на 478-й строке ", "", "");
                    }

                    //try
                    //{
                    //    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //    swApp.CloseDoc(modelName + ".sldprt");

                    //    if (makeDxf)
                    //    {
                    //        swApp.ExitApp();
                    //        swApp = null;
                    //    }
                    //    LoggerInfo("Обработка файла " + modelName + ".sldprt" + " успешно завершена", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    //    isErrors = false;
                    //}
                    //catch (Exception exception)
                    //{
                    //    LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    //    isErrors = true;
                    //    LoggerInfo("isErrors = true на 497-й строке ", "", "");
                    //}


                }

                #endregion


                try
                {
                    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    var namePrt = swApp.IActiveDoc2.GetTitle().ToLower().Contains(".sldprt")
                        ? swApp.IActiveDoc2.GetTitle()
                        : swApp.IActiveDoc2.GetTitle() + ".sldprt";
                    swApp.CloseDoc(namePrt);

                    if (makeDxf)
                    {
                        swApp.ExitApp();
                        swApp = null;
                    }
                    LoggerInfo(
                        "Обработка файла " + swApp.IActiveDoc2.GetTitle() + ".sldprt" + ".sldprt" + " успешно завершена",
                        "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = false;
                }
                catch (Exception exception)
                {
                    LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    isErrors = true;
                    LoggerInfo("isErrors = true на 497-й строке ", "", "");
                }
                finally
                {
                    //try
                    //{
                    //    swApp.CloseDoc(swApp.IActiveDoc2.GetTitle() + ".sldprt");
                    //}
                    //catch (Exception)
                    //{
                    //    swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //}
                }


            }
            catch (Exception exception)
            {
                LoggerError(String.Format("Общая ошибка метода: {0} Строка: {1} exception.Source - ", exception.Message, exception.StackTrace), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                if (swApp == null) return;
                if (makeDxf)
                {
                    swApp.ExitApp();
                }
                isErrors = true;
                LoggerInfo("isErrors = true на 506-й строке ", "", "");
            }
        }

        /// <summary>
        /// Creates the eprt.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="eDrwFileName">Name of the e DRW file.</param>
        /// <param name="isErrors">if set to <c>true</c> [is errors].</param>
        /// <param name="makeEprt">if set to <c>true</c> [make eprt].</param>
        public void CreateEprt(string filePath, out string eDrwFileName, out bool isErrors, bool makeEprt)
        {
            isErrors = false;
            eDrwFileName = "";
            
            SldWorks swApp = null;
            try
            {
                LoggerInfo("Запущен метод для обработки детали по пути " + filePath, "", "CreateFlattPatternUpdateCutlistAndEdrawing");

                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);

                try
                {
                    IEdmFolder5 oFolder;
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                    _currentVersion = edmFile5.CurrentVersion;
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при получении значения последней версии файла {0}", Path.GetFileName(filePath)), exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }
                
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = true };
                }
                if (swApp == null)
                {
                    isErrors = true;
                    return;
                }

                IModelDoc2 swModel;

                try
                {
                    swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    swModel.Extension.ViewDisplayRealView = false;
                    swModel.EditRebuild3();
                    swModel.ForceRebuild3(false);
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка при обработке детали {2}: {0} Строка: {1}", exception.Message, exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    return;
                }

                try
                {
                    if (!IsSheetMetalPart((IPartDoc)swModel))
                    {
                        swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                        swApp.ExitApp();
                        swApp = null;
                        return;
                    }
                }
                catch (Exception exception)
                {
                    LoggerError(String.Format("Ошибка2 при обработке детали {2}: {0} Строка: {1}", exception.Message, exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                }

                if (makeEprt)
                {
                    #region Сохранение детали в eDrawing

                    string modelName;
                    try
                    {
                        modelName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                        _eDrwFileName = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + modelName + ".eprt";
                        eDrwFileName = _eDrwFileName;
                    }
                    catch (Exception exception)
                    {
                        LoggerError(exception.ToString(), "", "");
                        isErrors = true;
                        LoggerInfo("isErrors = true на 423-й строке ", "", "");
                        return;
                    }

                    try
                    {
                        // todo: удаление документов перед новым сохранением. Осуществить поиск по имени
                        var existingDocument = SearchDoc(modelName + ".eprt", SwDocType.SwDocNone);

                        if (existingDocument != "")
                        {
                            LoggerInfo(String.Format("Файл есть в базе {0} и будет удален.. ", modelName), "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                            //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                            DeleteFileFromPdm(existingDocument, PdmBaseName);
                        }
                        else
                        {
                            File.Delete(_eDrwFileName);
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            File.Delete(_eDrwFileName);
                        }
                        catch (Exception exception)
                        {
                            LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                            isErrors = true;
                            LoggerInfo("isErrors = true на 453-й строке ", "", "");
                        }
                    }

                    #region ToDelete
                    //if (new FileInfo(_eDrwFileName).Exists)
                    //{
                    //    LoggerInfo("Файл есть в базе " + swModel.GetTitle());
                    //    //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                    //    DeleteFileFromPdm(_eDrwFileName, PdmBaseName);
                    //}
                    #endregion

                    try
                    {
                        swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption,
                        (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                        swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        swModel.Extension.SaveAs(_eDrwFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                    }
                    catch (Exception exception)
                    {
                        LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                        isErrors = true;
                    }

                    #endregion
                }

                try
                {
                    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    swApp.CloseDoc(swApp.IActiveDoc2.GetTitle() + ".sldprt");
                    LoggerInfo("Обработка файла " + swApp.IActiveDoc2.GetTitle() + ".sldprt" + ".sldprt" + " успешно завершена", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = false;
                }
                catch (Exception exception)
                {
                    LoggerError(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    isErrors = true;
                }


            }
            catch (Exception exception)
            {
                LoggerError(String.Format("Общая ошибка метода: {0} Строка: {1} exception.Source - ", exception.Message, exception.StackTrace), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                if (swApp == null) return;
                isErrors = true;
            }
        }


        /// <summary>
        /// Registrations in PDM.
        /// </summary>
        /// <param name="newFileName">Name of the eDRW file.</param>
        public bool RegistrationPdm(string newFileName)
        {
            return CheckInOutPdm(new List<FileInfo> { new FileInfo(newFileName) }, true, PdmBaseName);
        }

        #region Additional Methods

        static void DeleteFileFromPdm(string filePath, string pdmBase)
        {
            LoggerInfo(string.Format("Удаление файла по пути {0} базе PDM - {1}", filePath, pdmBase), "", "DeleteFileFromPdm");

            var retryCount = 2;
            var success = false;
            var ex = new Exception();
            while (!success && retryCount > 0)
            {
                try
                {
                    var vault1 = new EdmVault5();
                    IEdmFolder5 oFolder;
                    vault1.LoginAuto(pdmBase, 0);

                   // vault1.Login("kb81", "1", pdmBase);

                    vault1.GetFileFromPath(filePath, out oFolder);

                    var vault2 = (IEdmVault7)vault1;
                    var batchDeleter = (IEdmBatchDelete3)vault2.CreateUtility(EdmUtility.EdmUtil_BatchDelete);
                    batchDeleter.AddFileByPath(filePath);
                    batchDeleter.ComputePermissions(true);
                    batchDeleter.CommitDelete(0);
                    //LoggerInfo(string.Format("batchDeleter.CommitDelete - {0}", commitDelete));

                    LoggerInfo(string.Format("В базе PDM - {1}, удален файл по пути {0}", filePath, pdmBase), "", "DeleteFileFromPdm");

                    success = true;
                }
                catch (Exception exception)
                {
                    retryCount--;
                    ex = exception;
                    Thread.Sleep(200);
                    if (retryCount == 0)
                    {
                        // throw; //or handle error and break/return
                    }
                    LoggerError(String.Format("Во время удаления по пути {0} возникла ошибка. База - {1}. Ошибка: {2} Строка: {3}", filePath, pdmBase, exception.Message, exception.StackTrace), exception.Source, "DeleteFileFromPdm");
                }
            }
            if (!success)
            {
                LoggerError(String.Format("Во время удаления по пути {0} возникла ошибка. База - {1}. Ошибка: {2} Строка: {3}", filePath, pdmBase, ex.Message, ex.StackTrace), ex.Source, "DeleteFileFromPdm");
            }
        }

        static bool CheckInOutPdm(IEnumerable<FileInfo> filesList, bool registration, string pdmBase)
        {
            if (filesList.Count() == 1)
            {
                Thread.Sleep(5000);
            }
            foreach (var file in filesList)
            {
                var retryCount = 2;
                var success = false;
                var ex = new Exception();
                while (!success && retryCount > 0)
                {
                    try
                    {
                        var vault1 = new EdmVault5();
                        IEdmFolder5 oFolder;
                        vault1.LoginAuto(pdmBase, 0);

                        //vault1.Login("kb81","1",pdmBase);
                        
                        var edmFile5 = vault1.GetFileFromPath(file.FullName, out oFolder);
                        LoggerInfo(string.Format("Хранилище - {1}, файл {2} по пути {0}", file.FullName, pdmBase, edmFile5.Name), "", "CheckInOutPdm");
                        // Разрегистрировать
                        if (registration == false)
                        {
                            edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                            edmFile5.LockFile(oFolder.ID, 0);
                        }
                        // Зарегистрировать
                        if (registration)
                        {
                            Thread.Sleep(50);
                            edmFile5.UnlockFile(oFolder.ID, "");
                            Thread.Sleep(50);
                        }

                        LoggerInfo(string.Format("В хранилище - {1}, зарегестрирован документ по пути {0}", file.FullName, pdmBase), "", "CheckInOutPdm");

                        success = true;
                    }
                    catch (Exception exception)
                    {
                        retryCount--;
                        ex = exception;
                        Thread.Sleep(200);
                        if (retryCount == 0)
                        {
                            // throw; //or handle error and break/return
                        }
                        LoggerError(string.Format("Во время регистрации документа по пути {0} возникла ошибка{3}\nБаза - {1}. {2}", file.FullName, pdmBase, exception.Message, exception.StackTrace), exception.TargetSite.Name, "CheckInOutPdm");
                        return false;
                    }
                }
                if (success) continue;
                LoggerError(string.Format("Во время регистрации документа по пути {0} возникла ошибка\nБаза - {1}. {2}", file.FullName, pdmBase, ex.Message), ex.TargetSite.Name, "CheckInOutPdm");
                return false;
            }
            return true;
        }

        static void AddFileToPdm(string path, string pdmBase)
        {
            try
            {
                LoggerInfo(string.Format("Создание папки по пути {0} для сохранения", path), "", "AddFileToPdm");
                var vault1 = new EdmVault5();
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(pdmBase, 0);

                   // vault1.Login("kb81", "1", pdmBase);
                }

                var vault2 = (IEdmVault7)vault1;
                var fileDirectory = new FileInfo(path).DirectoryName;
                var fileFolder = vault2.GetFolderFromPath(fileDirectory);
                var result = fileFolder.AddFile(fileFolder.ID, "", Path.GetFileName(path));
                LoggerInfo(string.Format("Создание файла по пути {0} в папке {2} завершено. {1}", path, result, fileFolder.Name), "", "AddFileToPdm");
            }
            catch (Exception exception)
            {
                LoggerError(string.Format("Не удалось создать файл по пути {0}. Ошибка: {2} Строка: {1}", path, exception.StackTrace, exception.Message), exception.TargetSite.Name, "AddFileToPdm");
            }
        }

        void AddToPdmByPath(string path, string pdmBase)
        {
            try
            {
                //if (Directory.Exists(path))
                //{
                //    return;
                //}

                LoggerInfo(string.Format("Создание папки по пути {0} для сохранения", path), "", "AddToPdmByPath");
                var vault1 = new EdmVault5();
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(pdmBase, 0);
                }

                var vault2 = (IEdmVault7)vault1;
                //try
                //{
                //    var directoryInfo = new DirectoryInfo(path);
                //    if (directoryInfo.Parent == null) return;
                //    var parentFolder = vault2.GetFolderFromPath(directoryInfo.Parent.FullName);
                //    parentFolder.AddFolder(0, directoryInfo.Name);
                //    LoggerInfo(string.Format("Создание папки по пути {0} завершено.", path));
                //}
                //catch (Exception)
                //{
                    var fileDirectory = new FileInfo(path).DirectoryName;
                    var parentFolder = vault2.GetFolderFromPath(fileDirectory);

                
                   parentFolder.AddFile(parentFolder.ID, "",path);
                    //parentFolder.AddFolder(0, directoryInfo.Name);
                   LoggerInfo(string.Format("Создание файла по пути {0} завершено.", path), "", "AddToPdmByPath");
                //}
                
                
            }
            catch (Exception exception)
            {
                LoggerError(string.Format("Не удалось создать папку по пути {0}. Ошибка {1}", path, exception), "", "AddToPdmByPath");
            }
        }

        static void AddFilePdm(string path, string pdmBase)
        {
            try
            {
                //if (Directory.Exists(path))
                //{
                //    return;
                //}

                LoggerInfo(string.Format("Создание папки по пути {0} для сохранения", path), "", "AddToPdmByPath");
                var vault1 = new EdmVault5();
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(pdmBase, 0);
                }
                var vault2 = (IEdmVault7)vault1;
                var fileInfo = new FileInfo(path);
                if (fileInfo.Exists == false) return;
                var epdmFile = vault2.GetFolderFromPath(fileInfo.FullName);
            //    epdmFile.AddFolder(0, directoryInfo.Name);
              //  edmFolder6.AddFile()
            }
            catch (Exception exception)
            {
                LoggerError(string.Format("Не удалось создать папку по пути {0}. Ошибка {1}", path, exception), "", "AddToPdmByPath");
            }
        }
        
        static string GetFromCutlist(IModelDoc2 swModel, string property)
        {
            LoggerInfo(string.Format("Получение свойства '{1}' из CutList'а для {0}. Имя конфигурации '{2}'", new FileInfo(swModel.GetPathName()).Name, property, swModel.IGetActiveConfiguration().Name), "", "GetFromCutlist");

            var propertyValue = "";

            try
            {
                Feature swFeat2 = swModel.FirstFeature();
                while (swFeat2 != null)
                {
                    if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                    {
                        BodyFolder swBodyFolder = swFeat2.GetSpecificFeature2();
                        swFeat2.Select2(false, -1);
                        swBodyFolder.SetAutomaticCutList(true);
                        swBodyFolder.UpdateCutList();

                        Feature swSubFeat = swFeat2.GetFirstSubFeature();
                        while (swSubFeat != null)
                        {
                            if (swSubFeat.GetTypeName2() == "CutListFolder")
                            {
                                BodyFolder bodyFolder = swSubFeat.GetSpecificFeature2();
                                swSubFeat.Select2(false, -1);
                                bodyFolder.SetAutomaticCutList(true);
                                bodyFolder.UpdateCutList();
                                var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                //swCustPrpMgr.Add("Площадь поверхности", "Текст", "\"SW-SurfaceArea@@@Элемент списка вырезов1@ВНС-901.81.002.SLDPRT\"");
                                string valOut;
                                swCustPrpMgr.Get4(property, true, out valOut, out propertyValue);
                            }
                            swSubFeat = swSubFeat.GetNextFeature();
                        }
                    }
                    swFeat2 = swFeat2.GetNextFeature();
                }
                LoggerInfo("Метод GetFromCutlist() для " + new FileInfo(swModel.GetPathName()).Name + " завершен.", "", "GetFromCutlist");
            }
            catch (Exception exception)
            {
                LoggerError(string.Format("Во время получение свойства возникла ошибка: '{1}' в строке: {0}. Сообщение: '{2}'", exception.Source, exception.StackTrace, exception.Message), "", "GetFromCutlist");
            }
            
           
            return propertyValue;
        }

        static bool IsSheetMetalPart(IPartDoc swPart)
        {
            var mod = (IModelDoc2) swPart;

            LoggerInfo("Проверка на листовую деталь " + mod.GetTitle(), "", "IsSheetMetalPart");
            try
            {
                var isSheet = false;

                var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);

                foreach (Body2 vBody in vBodies)
                {
                    try
                    {
                        var isSheetMetal = vBody.IsSheetMetal();
                        if (!isSheetMetal) continue;
                        isSheet = true;
                    }
                    catch
                    {
                        isSheet = false;
                    }
                }

                LoggerInfo(String.Format("Проверка детали {0} завершена. Она {1}.", mod.GetTitle(),
                    isSheet ? "листовая" : "не листовая"), "", "IsSheetMetalPart");
                return isSheet;
            }
            catch (Exception)
            {
                LoggerInfo("Проверка завершена. Деталь не из листового материала.", "", "IsSheetMetalPart");
                // var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                // swApp.ExitApp();
                return false;
            }
        }

        #endregion


        //swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));


        #region Data to export

        class DataToExport
        {
            public string Config;
            public string Материал;
            public string Обозначение;
            public string ПлощадьПокрытия;
            public string КодМатериала;
            public string ДлинаГраничнойРамки;
            public string ШиринаГраничнойРамки;
            public string Сгибы;
            public string ТолщинаЛистовогоМеталла;
            public string Наименование;

            //public int Наименование;
        }
        
        void ExportDataToXmlSql(IModelDoc2 swModel, IEnumerable<DataToExport> dataToExport)
        {
            if (swModel == null || dataToExport == null)
            {
                // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                // ReSharper disable ConditionIsAlwaysTrueOrFalse
                LoggerError("Попытка запуска ExportDataToXmlSql() с пустыми параметрами. " + swModel == null ? "swModel = null" : "" + dataToExport == null ? "dataToExport = null" : "", "", "ExportDataToXmlSql");
                // ReSharper restore ConditionIsAlwaysTrueOrFalse
                return;
            }

            LoggerInfo("Выгрузка данных в XML файл и SQL базу по детали " + new FileInfo(swModel.GetPathName()).Name, "", "ExportDataToXmlSql");
            
            try
            {
                //var myXml = new System.Xml.XmlTextWriter(@"\\srvkb\SolidWorks Admin\XML\" + swModel.GetTitle() + ".xml", System.Text.Encoding.UTF8);
                //const string xmlPath = @"\\srvkb\SolidWorks Admin\XML\";
                //const string xmlPath = @"C:\Temp\";
                var myXml = new System.Xml.XmlTextWriter(_xmlPath + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".xml", System.Text.Encoding.UTF8);

                myXml.WriteStartDocument();
                myXml.Formatting = System.Xml.Formatting.Indented;
                myXml.Indentation = 2;

                // создаем элементы
                myXml.WriteStartElement("xml");
                myXml.WriteStartElement("transactions");
                myXml.WriteStartElement("transaction");

                myXml.WriteStartElement("document");

                foreach (var configData in dataToExport)
                {
                    #region XML

                    // Конфигурация
                    myXml.WriteStartElement("configuration");
                    myXml.WriteAttributeString("name", configData.Config);

                    // Материал
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Материал");
                    myXml.WriteAttributeString("value", configData.Материал);
                    myXml.WriteEndElement();

                    // Наименование  -- Из таблицы свойств
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Наименование");
                    myXml.WriteAttributeString("value", configData.Наименование);
                    myXml.WriteEndElement();

                    // Обозначение
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Обозначение");
                    myXml.WriteAttributeString("value", configData.Обозначение);
                    myXml.WriteEndElement();

                    // Площадь покрытия
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Площадь покрытия");
                    myXml.WriteAttributeString("value", configData.ПлощадьПокрытия);
                    myXml.WriteEndElement();

                    // ERP code
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Код_Материала");
                    myXml.WriteAttributeString("value", configData.КодМатериала);
                    myXml.WriteEndElement();

                    // Длина граничной рамки

                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Длина граничной рамки");
                    myXml.WriteAttributeString("value", configData.ДлинаГраничнойРамки);
                    myXml.WriteEndElement();

                    // Ширина граничной рамки
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Ширина граничной рамки");
                    myXml.WriteAttributeString("value", configData.ШиринаГраничнойРамки);
                    myXml.WriteEndElement();

                    // Сгибы
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Сгибы");
                    myXml.WriteAttributeString("value", configData.Сгибы);
                    myXml.WriteEndElement();

                    // Толщина листового металла
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Толщина листового металла");
                    myXml.WriteAttributeString("value", configData.ТолщинаЛистовогоМеталла);
                    myXml.WriteEndElement();

                    // Версия последняя
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Версия");
                    myXml.WriteAttributeString("value", Convert.ToString(_currentVersion));
                    myXml.WriteEndElement();

                    myXml.WriteEndElement();  //configuration

                    #endregion

                    #region SQL

                    try
                    {
                        // var sqlConnection = new SqlConnection(Settings.Default.SQLBaseCon);
                        //"Data Source=srvkb;Initial Catalog=SWPlusDB;Persist Security Info=True;User ID=sa;Password=PDMadmin;MultipleActiveResultSets=True");
                        var sqlConnection = new SqlConnection(_connectionString);
                        sqlConnection.Open();
                        var spcmd = new SqlCommand("UpDateCutList", sqlConnection) { CommandType = CommandType.StoredProcedure };
                        //spcmd.Parameters.Add("@MaterialsID", SqlDbType.Int).Value = КодМатериала;
                        var partNumber = configData.Обозначение;
                        var description = configData.Наименование;
                        double workpieceX; Double.TryParse(configData.ДлинаГраничнойРамки.Replace('.', ','), out workpieceX);
                            //Convert.ToDouble(configData.ДлинаГраничнойРамки.Replace('.', ','));
                        double workpieceY; Double.TryParse(configData.ШиринаГраничнойРамки.Replace('.', ','), out workpieceY);
                            //Convert.ToDouble(configData.ШиринаГраничнойРамки.Replace('.', ','));
                        int bend; Int32.TryParse(configData.ШиринаГраничнойРамки.Replace('.', ','), out bend);
                        //Convert.ToInt32(configData.Сгибы);
                        double thickness; Double.TryParse(configData.ТолщинаЛистовогоМеталла.Replace('.', ','), out thickness);
                            //(Double.TryParse(configData.ТолщинаЛистовогоМеталла.Replace('.', ','), out doubleValue))
                            //    ? doubleValue
                            //    : doubleValue; //Convert.ToDouble(configData.ТолщинаЛистовогоМеталла.Replace('.', ','));

                        var configuration = configData.Config;

                        spcmd.Parameters.Add("@PartNumber", SqlDbType.NVarChar).Value = partNumber;
                        spcmd.Parameters.Add("@Description", SqlDbType.NVarChar).Value = description;
                        if (configData.ДлинаГраничнойРамки == "") { workpieceX = 0; }
                        spcmd.Parameters.Add("@WorkpieceX", SqlDbType.Float).Value = workpieceX;
                        if (configData.ШиринаГраничнойРамки == "") { workpieceY = 0; }
                        spcmd.Parameters.Add("@WorkpieceY", SqlDbType.Float).Value = workpieceY;
                        if (configData.Сгибы == "") { bend = 0; }
                        spcmd.Parameters.Add("@Bend", SqlDbType.Int).Value = bend;
                        if (configData.ТолщинаЛистовогоМеталла == "") { thickness = 0; }
                        spcmd.Parameters.Add("@Thickness", SqlDbType.Float).Value = thickness;
                        spcmd.Parameters.Add("@Configuration", SqlDbType.NVarChar).Value = configuration;
                        spcmd.Parameters.Add("@version", SqlDbType.Int).Value = _currentVersion; //configData.versionPdm;
                        spcmd.ExecuteNonQuery();
                        sqlConnection.Close();
                    }
                    catch (Exception exception)
                    {
                        LoggerError(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.Message), exception.TargetSite.Name, "ExportDataToXmlSql");
                        MessageToUsr = exception.Message;
                    }

                    #endregion
                }

                //myXml.WriteEndElement();// ' элемент CONFIGURATION
                myXml.WriteEndElement();// ' элемент DOCUMENT
                myXml.WriteEndElement();// ' элемент TRANSACTION
                myXml.WriteEndElement();// ' элемент TRANSACTIONS
                myXml.WriteEndElement();// ' элемент XML
                // заносим данные в myMemoryStream
                myXml.Flush();

                myXml.Close();

                LoggerInfo("Выгрузка данных для детали " + swModel.GetTitle() + " завершена.", "", "ExportDataToXmlSql");
            }
            catch (Exception exception)
            {
                LoggerError(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.Message), exception.TargetSite.Name, "ExportDataToXmlSql");
                MessageToUsr = exception.Message;
            }
        }
        
        #endregion

        #region Save As Pdf

        /// <summary>
        /// Saves the DRW as PDF.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="pdfFilePath">The PDF file path.</param>
        public void SaveDrwAsPdf(string filePath, out string pdfFilePath)
        {
            pdfFilePath = "";
            
            //SearchDoc(Path.GetFileNameWithoutExtension(filePath), SwDocType.SwDocDrawing);
            
            SldWorks swApp = null;
            try
            {
                LoggerInfo(" Запущен метод сохранения чертежа для " + filePath, "", "SaveDrwAsPdf");

                var vault1 = new EdmVault5();
                IEdmFolder5 oFolder;
                vault1.LoginAuto(PdmBaseName, 0);

                var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = true };
                }
                if (swApp == null) { return; }

               
                var swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocDRAWING,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "00", 0, 0);
                swModel.Extension.ViewDisplayRealView = false;
                var swDraw = (DrawingDoc)swModel;
                
                try
                {
                    swDraw.ResolveOutOfDateLightWeightComponents();
                    swDraw.ForceRebuild();

                    // Движение по листам
                    var vSheetName = (string[]) swDraw.GetSheetNames();

                    foreach (var name in vSheetName)
                    {
                        swDraw.ResolveOutOfDateLightWeightComponents();
                        var swSheet = swDraw.Sheet[name];
                        swDraw.ActivateSheet(swSheet.GetName());

                        if ((swSheet.IsLoaded()))
                        {
                            try
                            {
                                var sheetviews = (object[]) swSheet.GetViews();
                                var firstView = (View) sheetviews[0];
                                firstView.SetLightweightToResolved();
                                
                                var baseView =  firstView.IGetBaseView();
                                var dispData = (IModelDoc2)baseView.ReferencedDocument;

                            }
                            catch (Exception exception)
                            {
                                LoggerError(exception.StackTrace + "\n", "", "SaveDrwAsPdf");
                               // MessageToUsr = exception.StackTrace;
                            }
                        }
                        else
                        {
                            return;
                        }

                        //Движение по видам
                        //if (!deep) continue;
                        try
                        {
                            var views = (object[]) swSheet.GetViews();
                            foreach (var drwView in views.Cast<View>())
                            {
                                drwView.SetLightweightToResolved();
                            }
                        }
                        catch (Exception exception)
                        {
                            LoggerError(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.Message), exception.TargetSite.Name, "SaveDrwAsPdf");
                        }
                    }
                    
                    #region Saving New Doc (Delete Old)

                    var errors = 0;
                    var warnings = 0;
                    var newpath = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".pdf";
                    pdfFilePath = newpath;

                    if (new FileInfo(newpath).Exists)
                    {
                        LoggerInfo("Файл есть в базе " + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".pdf", "", "SaveDrwAsPdf");
                        //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                        DeleteFileFromPdm(newpath, PdmBaseName);
                    }

                    swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swPDFExportInColor)), true);
                    swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swPDFExportEmbedFonts)), true);
                    swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights)), true);

                    var canSave = swModel.Extension.SaveAs(newpath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                    if (canSave)
                    {
                        DeleteFileFromPdm(newpath, PdmBaseName);
                        AddFileToPdm(newpath, PdmBaseName);
                        swModel.Extension.SaveAs(newpath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                    }

                    swApp.CloseDoc(Path.GetFileName(new FileInfo(newpath).FullName));

                    #endregion
                }
                catch (Exception exception)
                {
                    LoggerError(exception + "\n", "", "SaveDrwAsPdf");
                }
                finally
                {
                    swApp.ExitApp();
                    swApp = null;
                    LoggerInfo("PDF Сохранен", "", "SaveDrwAsPdf");
                }
            }
            catch (Exception exception)
            {
                if (swApp != null) swApp.ExitApp();
                LoggerError(exception + "\n", "", "SaveDrwAsPdf");
            }
        }

        #endregion

        #endregion

    }
}

#region To Delete

//Const EMPTY_DRAWING As String = "C:\EmptyDraw.SLDDRW"
//Const OUT_PATH As String = "\\srvkb\DXF"

//Dim swApp As SldWorks.SldWorks
//Dim swModel As SldWorks.ModelDoc2
//Dim swEmptyDraw As SldWorks.ModelDoc2
//Dim swPart As SldWorks.PartDoc

//Sub main()
                   
//    Dim fileName As String
//    fileName = "<Filepath>"
    
//    Dim ext As String
//    ext = UCase(Right(fileName, 6))
//    If ext <> "SLDPRT" Then
//        Exit Sub
//    End If
        
//    Set swApp = Application.SldWorks
    
//    swApp.Visible = True
//    Set swModel = swApp.OpenDoc6(fileName, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent + swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", 0, 0)
    
//    If swModel Is Nothing Then
//        Exit Sub
//    End If
    
//    If swModel.GetType <> swDocumentTypes_e.swDocPART Then
//        Exit Sub
//    End If
    
//    Set swPart = swModel
    
//    If Not IsSheetMetalPart() Then
//        Exit Sub
//    End If
                        
//    Set swEmptyDraw = swApp.OpenDoc6(EMPTY_DRAWING, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent + swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", 0, 0)
    
//    If swEmptyDraw Is Nothing Then
//        Exit Sub
//    End If

//    Dim vConfNames As Variant
//    Dim i As Integer
//    Dim outFile As String

//    vConfNames = swModel.GetConfigurationNames
    
//    For i = 0 To UBound(vConfNames)
        
//        Dim swConf As SldWorks.Configuration
//        Set swConf = swModel.GetConfigurationByName(vConfNames(i))
        
//        If False = swConf.IsDerived Then
            
//            swModel.ShowConfiguration2 vConfNames(i)
//            swModel.ForceRebuild3 False
//            outFile = GetOutFileName(fileName, CStr(vConfNames(i)))
    
//            If False = swEmptyDraw.Extension.SaveAs(outFile, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, 0, 0) Then
//                Exit Sub
//            End If
//            swPart.ExportFlatPatternView outFile, swExportFlatPatternViewOptions_e.swExportFlatPatternOption_RemoveBends
//        End If
    
//    Next
    
//    swApp.CloseDoc swEmptyDraw.GetTitle
//    swApp.CloseDoc swModel.GetTitle
 
//End Sub

//Function GetOutFileName(inputFile As String, confName As String) As String
    
//    Dim path As String
//    Dim name As String
    
//    path = Left(inputFile, InStrRev(inputFile, "\"))
//    name = Mid(inputFile, Len(path), Len(inputFile) - Len(path) - 6)
    
//    GetOutFileName = OUT_PATH + name + "-" + confName + ".dxf"
    
//End Function

//Function IsSheetMetalPart() As Boolean

//    Dim vBodies As Variant
//    Dim swBody As SldWorks.Body2
    
//    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, False)

//    Dim i As Integer
    
//    For i = 0 To UBound(vBodies)
        
//        Set swBody = vBodies(i)
        
//        If swBody.IsSheetMetal Then
//            IsSheetMetalPart = True
//            Exit Function
//        End If
        
//    Next
    
//    IsSheetMetalPart = False
    
//End Function

#endregion