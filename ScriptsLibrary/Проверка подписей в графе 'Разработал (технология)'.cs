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
    private const string signGraphValue = "Разработал (технология)";
    private List<int> TechIITypes;

    /// <summary>
    /// Заготовка
    /// </summary>
    const int zag = 1090;

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
                // объекты заготовок из извещения
                List<long> zagObjIDs = LoadItemIds(UserSession, techII, new List<int> { relTypeChangingByII, relTypeTechConsist, relTypeConsist },
                     new List<int>{zag}, -1);

                // проходим заготовки
                foreach (long zagObjId in zagObjIDs)
                {
                    IDBObject zagObj = UserSession.GetObject(zagObjId);

                    // подписи техпроцесса
                    List<long> signs = LoadItemIds(UserSession, zagObjId, new List<int> { relTypeChangingByII, relTypeSign },
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
                                zagObj.NameInMessages, user, signGraph);
                        }
                    }
                    // Проверка наличия подписи в графе signGraph
                    if (signs
                        .Select(sign => UserSession.GetObject(sign))
                        .Where(sign => sign.GetAttributeByGuid(new Guid("cad00141-306c-11d8-b4e9-00304f19f545")).Description == signGraphValue)
                        .ToList().Count == 0)
                    {
                        FinalMessage += string.Format("\r\nДля объекта {0} не задана ни одна подпись в графе {1}\r\n",
                            zagObj.NameInMessages, signGraphValue);
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