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

        string error = string.Empty;

        foreach (var attachment in activity.Attachments)
        {
            List<IDBObject> techProcesses = LoadItems(activity.Session, attachment.ObjectID,
                new List<int>() { MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)), 1, 1007 },
                new List<int>() { MetaDataHelper.GetObjectTypeID(new Guid("cad00187-306c-11d8-b4e9-00304f19f545" /*Техпроцесс единичный*/)) },
                -1);

            if (techProcesses == null)
                continue;

            ////количество повторений кода вида производства
            //Dictionary<long, int> workTypeDescCounts = new Dictionary<long, int>();

            foreach (var techProcess in techProcesses.OrderBy(e => e.Caption))
            {
                long workType = techProcess.GetAttributeByGuid(new Guid("cad0019c-306c-11d8-b4e9-00304f19f545" /*Вид производства*/)).AsInteger;

                if (workType == 0)
                    continue;

                IDBObject workTypeObj = activity.Session.GetObject(workType);

                long workTypeDescription = workTypeObj
                    .GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/))
                    .AsInteger;

                var arr = techProcess.Caption.Trim()
                    .Split(' ');
                var workTypeNo = Convert.ToInt32(
                    string.Join("", arr
                    [arr.Length - 2]
                    .Where(c => char.IsDigit(c))).ToString());

                //if (workTypeDescCounts.ContainsKey(workTypeDescription))
                //    workTypeDescCounts[workTypeDescription]++;
                //else
                //    workTypeDescCounts[workTypeDescription] = 1;

                List<IDBObject> techOperations = LoadItems(activity.Session, techProcess.ObjectID,
                    new List<int>() { MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)) },
                    MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/)),
                    -1);

                if (techOperations == null)
                    continue;

                foreach (var operation in techOperations)
                {
                    var operationNo = operation.GetAttributeByGuid(new Guid("cad009e6-306c-11d8-b4e9-00304f19f545" /*Номер объекта*/)).Value;
                    string operationNoInMO = string.Format("{0}.{1}.{2:000}", workTypeDescription, workTypeNo, operationNo);

                    var opAttr = operation.GetAttributeByGuid(new Guid("5a2e2fe6-d403-4249-b565-d372df44b803" /*Номер операции в сквозном МО*/));

					var normForOperList = LoadItems(activity.Session, operation.ObjectID,
                    new List<int>() { MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)) },
                    new List<int>() { MetaDataHelper.GetObjectTypeID(new Guid("cad005c2-306c-11d8-b4e9-00304f19f545" /*Нормирование на операцию*/)) },
                    -1);

                    if (normForOperList == null)
                        continue;

                    var normForOper = normForOperList.FirstOrDefault();

                    var normObjAttr = normForOper != null ? normForOper.GetAttributeByGuid(new Guid("cadd93a5-306c-11d8-b4e9-00304f19f545" /*Штучно-калькуляционное время*/)) : null;

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(normObjAttr != null ? normObjAttr.AsString : string.Empty))
                            opAttr.Value = operationNoInMO;
                    }
                    catch (Exception exc)
                    {
                        error += exc.Message + "\n";
                    }
                }
            }
        }

        if (error.Length > 0)
            throw new NotificationException(error);
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