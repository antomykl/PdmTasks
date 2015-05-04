using System;
using EdmLib;

namespace MakeDxfUpdatePartData
{
    /// <summary>
    /// Основной класс библиотеки
    /// </summary>
    public partial class MakeDxfExportPartDataClass
    {
        private IEdmVault5 _edmVault5;

        /// <summary>
        /// Тип документа для поиска. 
        /// </summary>
        public enum SwDocType
        {
            /// <summary>
            /// Точное имя файла для поиска
            /// </summary>
            SwDocNone = 0,
            /// <summary>
            /// Имя файла с расширением .sldprt
            /// </summary>
            SwDocPart = 1,
            /// <summary>
            /// Имя файла с расширением .sldasm
            /// </summary>
            SwDocAssembly = 2,
            /// <summary>
            /// Имя файла с расширением .slddrw
            /// </summary>
            SwDocDrawing = 3,
            /// <summary>
            /// По схожести
            /// </summary>
            SwDocLike = 4,
        }

        /// <summary>
        /// Searches the document.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="swDocType">Type of the sw document.</param>
        /// <returns></returns>
        public string SearchDoc(string filePath, SwDocType swDocType)
        {
            try
            {
                if (_edmVault5 == null)
                {
                    _edmVault5 = new EdmVault5();
                }
                var edmVault7 = (IEdmVault7)_edmVault5;

                if (!_edmVault5.IsLoggedIn)
                {
                    _edmVault5.LoginAuto(PdmBaseName, 0);
                }

                //Search for all text files in the edmVault7
                var edmSearch5 = (IEdmSearch5)edmVault7.CreateUtility(EdmUtility.EdmUtil_Search);
                
                string  extenison;
                string like;

                switch ((int)swDocType)
                {
                    case 0:
                        extenison = "";
                        like = "";
                        break;
                    case 1:
                        like = "";
                        extenison = ".sldprt";
                        break;
                    case 2:
                        like = "";
                        extenison = ".sldasm";
                        break;
                    case 3:
                        like = "%";
                        extenison = ".slddrw";
                        break;
                    case 4:
                        like = "%";
                        extenison = ".*";
                        break;
                    default:
                        extenison = "";
                        like = "";
                        break;
                }
                edmSearch5.FileName = like + filePath + extenison;

                var edmSearchResult5 = edmSearch5.GetFirstResult();

                if (edmSearch5.GetFirstResult() == null)
                {
                    LoggerInfo(String.Format("Файл с именем {0} не найден!", like + filePath + extenison), "", "SearchDoc");
                    return "";
                }

                IEdmFolder5 edmFolder5;
                var edmFile5 = _edmVault5.GetFileFromPath(edmSearchResult5.Path, out edmFolder5);
                ShowReferences((EdmVault5)_edmVault5, edmSearchResult5.Path);
                edmFile5.GetFileCopy(0, "", edmSearchResult5.Path);
                filePath = edmSearchResult5.Path;
            }

            catch (Exception exception)
            {
                LoggerError(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.Message), exception.TargetSite.Name, "SearchDoc");
            }

            return filePath;
        }

        void ShowReferences(IEdmVault7 edmVault7, string filePath)
        {
            string projName = null;
            IEdmFolder5 edmFolder5;
            var edmFile5 = edmVault7.GetFileFromPath(filePath, out edmFolder5);
            var edmReference5 = edmFile5.GetReferenceTree(edmFolder5.ID);
            AddReferences(edmReference5, 0, ref projName, false);
        }

        string AddReferences(IEdmReference5 file, long indent, ref string projName, bool isTop)
        {
            var fileName = file.Name;

            if (_edmVault5 == null)
            {
                _edmVault5 = new EdmVault5();
            }
            if (!_edmVault5.IsLoggedIn)
            {
                _edmVault5.LoginAuto(PdmBaseName, 0);
            }
            var edmPos5 = file.GetFirstChildPosition(projName, isTop, true);
            while (!(edmPos5.IsNull))
            {
                var edmReference5 = file.GetNextChild(edmPos5);
                IEdmFolder5 edmFolder5;
                var edmFile5 = _edmVault5.GetFileFromPath(edmReference5.FoundPath, out edmFolder5);
                fileName = fileName + AddReferences(edmReference5, indent, ref projName, isTop);
                edmFile5.GetFileCopy(0, "", edmReference5.FoundPath);
            }
            return fileName;
        }
    }
}
