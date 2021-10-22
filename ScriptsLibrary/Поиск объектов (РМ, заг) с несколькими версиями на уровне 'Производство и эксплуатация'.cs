//Поиск объектов (дет и СЕ) с несколькими версиями на уровне 'Производство и эксплуатация'

// находит в сборке все подсборки и детали, проверяет у каждого объекта наличие нескольких версий на уровне 'Производство и эксплуатация' 
// и гененирует исключение в виде списка ошибок
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; set; }

    #region Константы

    private const int inProductionLCstep = 1058;

    #endregion Константы

    #region Типы связей

    private const int relTypeConsist = 1; //Состав изделия

    #endregion Типы связей

    public void Execute(IActivity activity)
    {
        if (Debugger.IsAttached)
        {
            Debugger.Break();
        }

        string resultMessage = string.Empty;
        IUserSession session = activity.Session;

        //только одно вложение
        IAttachment attachment = activity.Attachments[0];

        List<long> allObjs = LoadItemIds(session, attachment.ObjectID, new List<int> { relTypeConsist, MetaDataHelper.GetRelationTypeID(new System.Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/))  }, new List<int> { 1090, 1176, 1292  }, -1);
        allObjs.Add(attachment.ObjectID);
        for (int i = 0; i < allObjs.Count; i++)
        {
            long item = allObjs[i];
            if (session.GetObject(item).GetAttributeByGuid(new System.Guid("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/)).AsString != "БМЗ")
                continue;

            List<long> allObjVersions = session.GetObjectIDVersions(item);
            if (allObjVersions.Count > 1)
            {
                List<IDBObject> prodVersions = allObjVersions
                    .Select(version => session.GetObject(version))
                    .Where(version => version.LCStep == inProductionLCstep)
                    .ToList();

                if (prodVersions.Count > 1)
                {
                    string wrongObjsStr = string.Empty;
                    prodVersions.ForEach(obj =>
                    {
                        wrongObjsStr += obj.ObjectID + "; ";
                    });
                    resultMessage += string.Format("\n\rУ объекта {0}({1}) обнаружены {2} версии на шаге 'Производство и эксплуатация': \n {3}", session.GetObject(item).NameInMessages, item, prodVersions.Count, wrongObjsStr);
                }
            }
        }
        if (resultMessage.Length > 0)
        {
            throw new NotificationException(resultMessage);
        }
    }

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
}