//Заполнение номера операции в сквозном МО
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    public void Execute(IActivity activity)
    {
        if (System.Diagnostics.Debugger.IsAttached)
            System.Diagnostics.Debugger.Break();

        foreach (var attachment in activity.Attachments)
        {
            List<IDBObject> techProcesses = LoadItems(activity.Session, attachment.ObjectID,
                new List<int>() { MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)), 1 },
                MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00185-306c-11d8-b4e9-00304f19f545" /*Техпроцесс базовый*/)),
                -1);

            if (techProcesses == null)
                continue;

            //количество повторений кода вида производства
            Dictionary<long, int> workTypeDescCounts = new Dictionary<long, int>();
            foreach (var techProcess in techProcesses.OrderBy(e => e.Caption))
            {
                long workType = techProcess.GetAttributeByGuid(new Guid("cad0019c-306c-11d8-b4e9-00304f19f545" /*Вид производства*/)).AsInteger;
                long workTypeDescription = activity.Session
                    .GetObject(workType)
                    .GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/))
                    .AsInteger;

                if (workTypeDescCounts.ContainsKey(workTypeDescription))
                    workTypeDescCounts[workTypeDescription]++;
                else
                    workTypeDescCounts[workTypeDescription] = 1;

                List<IDBObject> techOperations = LoadItems(activity.Session, techProcess.ObjectID,
                    new List<int>() { MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)) },
                    MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/)),
                    -1);

                if (techOperations == null)
                    continue;

                foreach (var operation in techOperations)
                { 
                    var operationNo = operation.GetAttributeByGuid(new Guid("cad009e6-306c-11d8-b4e9-00304f19f545" /*Номер объекта*/)).Value;

                    string operationNoInMO = string.Format("{0}.{1}.{2:000}", workTypeDescription, workTypeDescCounts[workTypeDescription], operationNo);

                    operation.GetAttributeByGuid(new Guid("a6b34136-17d9-480e-8658-29642e557591" /*Номер операции для ERP*/)).Value = operationNoInMO;
                }
            }
        }
    }

    private List<IDBObject> LoadItems(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
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
                .Select(element => session.GetObject((long)element[0]))
                .ToList();
    }

}