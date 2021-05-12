using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Contexts;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using System.Diagnostics;
using Intermech.Interfaces.Client;
public class Script
{
    public ICSharpScriptContext ScriptContext { get; set; }

    #region Константы
    /// <summary>
    /// Графа подписи 
    /// </summary>
    private const string signGraphValue = "Врем.нормир.";
    private List<int> TechIITypes;

    /// <summary>
    /// ТП единичный
    /// </summary>
    const int oneTP = 1237;

    /// <summary>
    /// нормирование на операцию, переход и техпроцесс
    /// </summary>
    private List<int> objNorm;

    private const int TechIIType = 2761;
    private const int TechSvIIType = 2762;
    private const int signType = 1025;
    private const int criptSignType = 1101;

    #endregion Константы

    #region Типы связей

    private const int relTypeChangingByII = 1007;//Изменяется по извещению
    private const int relTypeConsist = 1; //Состав изделия
    private const int relTypeTechConsist = 1002; //Технологический состав
    private const int relTypeSign = 1013; //Состав подписей

    #endregion Типы связей

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
                // техпроцессы из извещения
                List<long> techprocesses = LoadItemIds(UserSession, techII, new List<int> { relTypeChangingByII, relTypeTechConsist, relTypeConsist },
                    new List<int> { oneTP, /*ТП групповой = */ 1270, /*ТП типовой = */ 1255 }, -1);

                // объекты нормирования из извещения
                List<long> normObjs = LoadItemIds(UserSession, techII, new List<int> { relTypeChangingByII, relTypeTechConsist, relTypeConsist },
                     objNorm, -1);

                // проходим техпроцессы
                foreach (long techprocess in techprocesses)
                {

                    IDBObject techprocessObj = UserSession.GetObject(techprocess);

                    //ТП на производстве
                    if (techprocessObj.LCStep == 1058)
                        continue;

                    // подписи техпроцесса
                    List<long> signs = LoadItemIds(UserSession, techprocess, new List<int> { relTypeChangingByII, relTypeSign },
                      new List<int> { signType, criptSignType }, 1);

                    foreach (long sign in signs)
                    {
                        IDBObject signObj = UserSession.GetObject(sign);

                        DateTime modifyDate = signObj.GetAttributeByGuid(new Guid("cad0013a-306c-11d8-b4e9-00304f19f545"
                            /*Дата модификации содержимого объекта*/)).AsDateTime;

                        DateTime signDate = signObj.GetAttributeByGuid(new Guid("cad014cb-306c-11d8-b4e9-00304f19f545"
                            /*Дата подписания*/)).AsDateTime;

                        string signGraph = signObj.GetAttributeByGuid(new Guid("cad00141-306c-11d8-b4e9-00304f19f545")).Description;
                        string user = signObj.GetAttributeByGuid(new Guid("cad00143-306c-11d8-b4e9-00304f19f545")).Description;

                        if (modifyDate > signDate && signGraph == signGraphValue)
                        {
                            FinalMessage += string.Format("\r\nДля {0} требуется обновить подпись пользователя {1} в графе {2}.\r\n",
                                techprocessObj.NameInMessages, user, signGraph);
                        }
                    }
                    // Проверка наличия подписи в графе signGraph
                    if (signs
                        .Select(sign => UserSession.GetObject(sign))
                        .Where(sign => sign.GetAttributeByGuid(new Guid("cad00141-306c-11d8-b4e9-00304f19f545")).Description == signGraphValue)
                        .ToList().Count == 0)
                    {
                        FinalMessage += string.Format("\r\nДля {0} не задана ни одна подпись в графе {1}\r\n",
                            techprocessObj.NameInMessages, signGraphValue);
                    }

                }

                // проходим объекты нормирования
                foreach (long normObj in normObjs)
                {
                    IDBObject normObject = UserSession.GetObject(normObj);

                    if (normObject.CheckoutBy > 0)
                    {
                        IDBObject redactor = UserSession.GetObject(normObject.CheckoutBy);

                        FinalMessage += string.Format("\r\n[{0}] взял объект [{1}]|[{2}] на редактирование.\r\n",
                        redactor.NameInMessages, normObject.NameInMessages, normObject.ObjectID);
                    }
                }
            }
        }

        if (FinalMessage.Length > 0)
        {
            throw new Exception(FinalMessage);
        }
    }
}