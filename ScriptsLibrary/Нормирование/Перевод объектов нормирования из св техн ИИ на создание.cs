using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Contexts;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using Intermech.Workflow;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scripts
{
    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; set; }

        #region Константы

        /// <summary>
        /// ТП единичный
        /// </summary>
        const int oneTP = 1237;

        List<int> TechIITypes;

        /// <summary>
        /// нормирование на операцию, переход и техпроцесс
        /// </summary>
        List<int> objNorm;

        const int TechIIType = 2761;
        const int TechSvIIType = 2762;

        const int lcStep_Sign = 1056;
        const int lcStep_Development = 1050;
        #endregion

        #region Типы связей
        const int relTypeChangingByII = 1007;//Изменяется по извещению
        const int relTypeConsist = 1; //Состав изделия
        const int relTypeTechConsist = 1002; //Технологический состав

        #endregion

        #region ID Организаций-источников

        const long idKSK = 63822858; // id объекта организации ОП "КСК-Брянск"


        #endregion

        /// <summary>
        /// Получение списка ID состава/применяемости
        /// </summary>
        /// <param name="session"></param>
        /// <param name="objID"></param>
        /// <param name="relIDs"></param>
        /// <param name="childObjIDs"></param>
        /// <param name="lv">Количество уровней для обработки (-1 = загрузка полного состава / применяемости)</param>
        /// <returns></returns>

        private List<long> LoadItemIds(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
        {
            ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
            List<ColumnDescriptor> col = new List<ColumnDescriptor>
            {
                new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
            };
            List<long> reslt = new List<long>();
            session.EditingContextID = objID;
            DataTable dt = _loadService.LoadComplexCompositions(
            session.SessionGUID,
            new List<ObjInfoItem> { new ObjInfoItem(objID) },
            relIDs,
            childObjIDs,
            col,
            true, // Режим получения данных (true - состав, false - применяемость)
            false, // Флаг группировки данных
            null, // Правило подбора версий
            null, // Условия на объекты при получении состава / применяемости
                  //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
                  //SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
            Intermech.SystemGUIDs.filtrationLatestVersions,
            null,
            lv // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
            );
            if (dt == null)
                return reslt;
            else
                return dt.Rows.OfType<DataRow>()
            .Select(element => (long)element[0]).ToList<long>();
        }

        /// <summary>
        /// Проверка атрибута "организация-источник БМЗ"
        /// </summary>
        private bool? FromBMZ(IUserSession UserSession, long elem)
        {
            const long idBMZ = 3251710; // id ообъекта организации "БМЗ"
            IDBObject obj = UserSession.GetObject(elem);
            IDBAttribute orgIstAtr = obj.GetAttributeByName("Организация-источник");
            if (orgIstAtr.IsNull)
            {
                return null;
            }
            long org = (long)orgIstAtr.Value;
            if (org == idBMZ)
                return true;
            else
                return false;
        }

        public void Execute(IActivity activity)
        {
            IUserSession UserSession = activity.Session;
            string FinalMessage = string.Empty;
            string mailToAdmins = string.Empty;

            objNorm = new List<int> { 1212, 1193, 1175 };

            TechIITypes = new List<int>();
            TechIITypes.Add(TechIIType);
            TechIITypes.Add(TechSvIIType);

            foreach (IAttachment attachment in activity.Attachments)
            {
                IDBEditingContextsService contextService = UserSession.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;
                EditingContextsObjectContainer contextContainer = contextService.GetEditingContextsObject(UserSession.SessionGUID, attachment.ObjectID, true, true);

                List<long> techIIs = contextContainer.GetContextsID()
                    .Where(ii => TechIITypes.Contains(UserSession.GetObject(ii).TypeID))
                    .ToList();

                foreach (long techII in techIIs)
                {
                    // объекты нормирования из извещения
                    List<long> normObjs = LoadItemIds(UserSession, techII, new List<int> { relTypeChangingByII, relTypeTechConsist, relTypeConsist },
                         objNorm, -1);

                    // проходим объекты нормирования
                    foreach (long normObj in normObjs)
                    {
                        IDBObject normObject = UserSession.GetObject(normObj);

                        if (FromBMZ(UserSession, normObj) == true)
                        {
                            if (normObject.LCStep == lcStep_Sign) 
                            {
                                try
                                {
                                    normObject.LCStep = lcStep_Development;
                                }
                                catch (KernelException exc)
                                {
                                    FinalMessage += exc.Message + "\n\r" + exc.StackTrace;
                                }
                            }
                        }
                        else
                        {
                            FinalMessage += string.Format("Объект {0} (id{1}) не принадлежит БМЗ.\n\r", normObject.NameInMessages, normObj);
                        }
                    }                 
                }
            }

            if (FinalMessage.Length > 0)
            {
                throw new NotificationException(FinalMessage);
            }
        }
    }
}
