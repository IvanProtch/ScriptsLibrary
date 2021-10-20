//Заполнение номера операции в сквозном МО
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    public class Item
    {
        private Dictionary<long, Item> _entersInAsms = new Dictionary<long, Item>();

        public Item() { }

        public Item(DataRow row, Dictionary<long, Item> itemsDictionary)
        {
            Parents = new List<Item>();
            Childs = new List<Item>();

            ObjectId = Convert.ToInt64(row["F_OBJECT_ID"]);
            ObjectType = Convert.ToInt32(row["F_OBJECT_TYPE"]);
            Caption = Convert.ToString(row["CAPTION"]);

            long parentId = Convert.ToInt64(row["F_PROJ_ID"]);

            if (itemsDictionary.ContainsKey(parentId))
            {
                Item parent = itemsDictionary[parentId];
                parent.Childs.Add(this);
                this.Parents.Add(parent);
            }

            itemsDictionary[this.ObjectId] = this;
        }

        public string Caption { get; set; }
        public long ObjectId { get; internal set; }
        public int ObjectType { get; internal set; }

        public List<Item> Parents { get; set; }
        public List<Item> Childs { get; set; }

        public Dictionary<long, Item> EntersIn
        {
            get
            {
                if (_entersInAsms.Keys.Count == 0)
                {
                    foreach (var parent in Parents)
                    {
                        if (MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid(SystemGUIDs.objtypeProduct)).Contains(parent.ObjectType))
                        {
                            if (!_entersInAsms.ContainsKey(parent.ObjectId))
                                _entersInAsms[parent.ObjectId] = parent;
                        }
                        else
                        {
                            foreach (var nextAsm in parent.EntersIn)
                            {
                                if (!_entersInAsms.ContainsKey(nextAsm.Key))
                                    _entersInAsms[nextAsm.Key] = nextAsm.Value;
                            }
                        }
                    }
                }
                return _entersInAsms;
            }
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, childs={3}, parents={4}", ObjectId, Caption, MetaDataHelper.GetObjectTypeName(ObjectType), Childs.Count, Parents.Count);
        }
    }

    public class TechProcess : Item
    {
        private Dictionary<long, string> _productionTypeTable = new Dictionary<long, string>();
        public TechProcess() { }
        public TechProcess(DataRow row, Dictionary<long, Item> itemsDictionary, Dictionary<long, string> productionTypeTable) : base(row, itemsDictionary)
        {
            _productionTypeTable = productionTypeTable;
            var data = row["cad0019c-306c-11d8-b4e9-00304f19f545" /*Вид производства*/];
            if (data is long)
                WorkTypeDescription = _productionTypeTable[Convert.ToInt64(data)];
        }
        public string WorkTypeDescription { get; set; }
        private int _workTypeIndex = 0;
        public int WorkTypeIndex 
        {
            get
            {
                if(_workTypeIndex == 0)
                {
                    var arr = this.Caption.Trim()
                         .Split(' ');
                    
                    _workTypeIndex = Convert.ToInt32(
                        string.Join("", arr
                        [arr.Length - 2]
                        .Where(c => char.IsDigit(c))).ToString());
                }
                return _workTypeIndex;
            }
        }

        public void CalculateOperationsNoInMO()
        {
            foreach (var workShop in this.Childs)
            {
                foreach (Operation operation in workShop.Childs)
                {
                    NormForOperation norm = operation.Childs.FirstOrDefault() as NormForOperation;
                    if (norm == null)
                        continue;
                    else
                    {
                        if (norm.PartTime == null)
                            continue;
                    }

                    operation.OpNoInMO = string.Format("{0}.{1}.{2:000}", this.WorkTypeDescription, this.WorkTypeIndex, operation.No);
                }
            }
        }
    }
    public class NormForOperation : Item
    {
        public NormForOperation() { }
        public NormForOperation(DataRow row, Dictionary<long, Item> itemsDictionary) : base(row, itemsDictionary)
        {
            var data = row["cadd93a5-306c-11d8-b4e9-00304f19f545" /*Штучно-калькуляционное время*/];
            if (data is long)
                PartTime = Convert.ToString(data);
        }

        public string PartTime { get; set; }


    }
    public class Operation : Item
    {
        public Operation() { }
        public Operation(DataRow row, Dictionary<long, Item> itemsDictionary) : base(row, itemsDictionary)
        {
            var value = row["cad009e6-306c-11d8-b4e9-00304f19f545" /*Номер объекта*/];
            if (value is int)
                No = Convert.ToInt32(value);
            else
                No = 0;

            OpNoInMO = Convert.ToString(row["5a2e2fe6-d403-4249-b565-d372df44b803" /*Номер операции в сквозном МО*/]);
        }
        public int No { get; set; }
        public string OpNoInMO { get; set; }

        public override string ToString()
        {
            return string.Format(base.ToString() + ", {0}", OpNoInMO);
        }
    }

    public void Execute(IActivity activity)
    {
        if (System.Diagnostics.Debugger.IsAttached)
            System.Diagnostics.Debugger.Break();

        string error = string.Empty;

        foreach (var attachment in activity.Attachments)
        {
            Dictionary<long, Item> techProcesses = new Dictionary<long, Item>();

            LoadComposition(activity.Session, new List<long>() { attachment.ObjectID }, techProcesses);

            if (techProcesses == null)
                continue;

            foreach (TechProcess techProcess in techProcesses.Values.OfType<TechProcess>().OrderBy(e => e.Caption))
            {
                techProcess.CalculateOperationsNoInMO();
            }

            foreach (Operation operation in techProcesses.Values.OfType<Operation>())
            {
                if (operation.OpNoInMO == null || operation.OpNoInMO == string.Empty)
                    continue;

                var attr = activity.Session.GetObject(operation.ObjectId).GetAttributeByGuid(new Guid("5a2e2fe6-d403-4249-b565-d372df44b803" /*Номер операции в сквозном МО*/));

                try
                {
                    if (attr != null)
                        attr.Value = operation.OpNoInMO;
                }
                catch (Exception exc)
                {
                    error += exc.Message + "\n";
                }
            }
        }
        if (error.Length > 0)
            throw new NotificationException(error);
    }

    private void LoadComposition(IUserSession session, List<long> ids, Dictionary<long, Item> itemsDictionary)
    {
        Dictionary<long, Item> result = itemsDictionary;

        IDBObjectCollection prodType = session.GetObjectCollection(MetaDataHelper.GetObjectTypeID(new Guid("cad005ae-306c-11d8-b4e9-00304f19f545")));
        ColumnDescriptor[] columnsT = new ColumnDescriptor[]
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
            new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID(new Guid("cad0001f-306c-11d8-b4e9-00304f19f545" /*Обозначение*/)), AttributeSourceTypes.Object, ColumnContents.String, ColumnNameMapping.Guid, SortOrders.NONE, 0),
        };
        DBRecordSetParams dBRecord = new DBRecordSetParams(new ConditionStructure[] { }, columnsT);
        DataTable dt = prodType.Select(dBRecord);
        Dictionary<long, string> productionTypeTable = new Dictionary<long, string>();
        foreach (DataRow type in dt.Rows)
        {
            productionTypeTable[Convert.ToInt64(type["F_OBJECT_ID"])] = Convert.ToString(type["cad0001f-306c-11d8-b4e9-00304f19f545" /*Обозначение*/]);
        }

        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;

        List<int> relIDs = new List<int>()
        {
            MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)),
            MetaDataHelper.GetRelationTypeID("cad00023-306c-11d8-b4e9-00304f19f545" /*Состоит из*/),
            MetaDataHelper.GetRelationTypeID("cad0036b-306c-11d8-b4e9-00304f19f545" /*Изменяется по извещению*/)
        };

        List<int> childObjIDs = new List<int>()
        {
            MetaDataHelper.GetObjectTypeID(new Guid("cad00187-306c-11d8-b4e9-00304f19f545" /*Техпроцесс единичный*/)),
            MetaDataHelper.GetObjectTypeID(new Guid("cad005c2-306c-11d8-b4e9-00304f19f545" /*Нормирование на операцию*/)),
            MetaDataHelper.GetObjectTypeID(new Guid("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/)),
            MetaDataHelper.GetObjectTypeID(new Guid("cad001ff-306c-11d8-b4e9-00304f19f545" /*Цехозаход*/)),
            MetaDataHelper.GetObjectTypeID("cad0016f-306c-11d8-b4e9-00304f19f545" /*Маршрут обработки*/),
            MetaDataHelper.GetObjectTypeID("cad00132-306c-11d8-b4e9-00304f19f545" /*Сборочные единицы*/)
        };

        childObjIDs.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)));

        List<ColumnDescriptor> columns = new List<ColumnDescriptor>
        {
             new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.CAPTION, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_PROJ_ID, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0),

            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_TYPE, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0),

            new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID(new Guid("cad0019c-306c-11d8-b4e9-00304f19f545" /*Вид производства*/)), AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Guid, SortOrders.NONE, 0),

            new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID("cadd93a5-306c-11d8-b4e9-00304f19f545" /*Штучно-калькуляционное время*/), AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0),

            new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID(new Guid("cad009e6-306c-11d8-b4e9-00304f19f545" /*Номер объекта*/)),
            AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0),

            new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID(new Guid("5a2e2fe6-d403-4249-b565-d372df44b803" /*Номер операции в сквозном МО*/)),
            AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0)
        };

        List<ConditionStructure> conditions = new List<ConditionStructure>()
        {
            new ConditionStructure(new Guid("cadd93a5-306c-11d8-b4e9-00304f19f545" /*Штучно-калькуляционное время*/), RelationalOperators.NotEmpty, 0, LogicalOperators.NONE, 0),

            //new ConditionStructure(new Guid("cad0019c-306c-11d8-b4e9-00304f19f545" /*Вид производства*/), RelationalOperators.NotEqual, 0, LogicalOperators.NONE, 0),
        };

        dt = _loadService.LoadComplexCompositions(
        session.SessionGUID,
        ids.Select(e => new ObjInfoItem(e)).ToList(),
        relIDs,
        childObjIDs,
        columns,
        true, // Режим получения данных (true - состав, false - применяемость)
        false, // Флаг группировки данных
        null, // Правило подбора версий
        null, // Условия на объекты при получении состава / применяемости
              //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
              //SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
        Intermech.SystemGUIDs.filtrationLatestVersions,
        null,
        -1 // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
        );

        int norm = MetaDataHelper.GetObjectTypeID(new Guid("cad005c2-306c-11d8-b4e9-00304f19f545" /*Нормирование на операцию*/));
        int op = MetaDataHelper.GetObjectTypeID(new Guid("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/));
        int tp = MetaDataHelper.GetObjectTypeID(new Guid("cad00187-306c-11d8-b4e9-00304f19f545" /*Техпроцесс единичный*/));
        foreach (DataRow row in dt.Rows)
        {
            int type = Convert.ToInt32(row["F_OBJECT_TYPE"]);

            if (type == op)
                new Operation(row, result);
            else if (type == tp)
                new TechProcess(row, result, productionTypeTable);
            else if (type == norm)
                new NormForOperation(row, result);
            else
                new Item(row, result);
        }

    }

}