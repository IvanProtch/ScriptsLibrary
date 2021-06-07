//Проверка редактируемости состава объекта

// Проверяет состав любого объекта, если удается найти объект взятый на редактирование, формирует сообщение об ошибке.
using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Workflow;
using System.Xml.Linq;
using System.Xml;
using System.Diagnostics;
using Intermech.Interfaces.Compositions;
using System.Data;
using Intermech.Kernel.Search;
using System.Linq;
using System.Collections.Generic;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    public void Execute(IActivity activity)
    {
        List<int> allRelations = MetaDataHelper.GetRelationTypesList().Select(rel => rel.RelationTypeID).ToList();
        string error = string.Empty;
        foreach (IAttachment attachment in activity.Attachments)
        {
            List<IDBObject> consistance = GetObjectConsistance(activity.Session, attachment.ObjectID, allRelations, null, -1);
            foreach (IDBObject item in consistance)
            {
                if (item.CheckoutBy > 0)
                    error += string.Format("У объекта '{0}' не может быть изменен шаг жизненного цикла, пока объект '{1}' находится на редактировании.\r\n", attachment.Object.NameInMessages, item.NameInMessages);
            }
        }
        if (error.Length > 0)
            throw new NotificationException(error);
    }

    /// <summary>
    /// Получение списка состава/применяемости
    /// </summary>
    /// <param name="session"></param>
    /// <param name="objID"></param>
    /// <param name="relIDs"></param>
    /// <param name="childObjIDs"></param>
    /// <param name="lv">Количество уровней для обработки (-1 = загрузка полного состава / применяемости)</param>
    /// <returns></returns>
    private List<IDBObject> GetObjectConsistance(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
    {
        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
        List<ColumnDescriptor> col = new List<ColumnDescriptor>
            {
                new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
            };
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
            return null;
        else
            return dt.Rows.OfType<DataRow>()
        .Select(element => session.GetObject((long)element[0])).ToList();
    }
}
