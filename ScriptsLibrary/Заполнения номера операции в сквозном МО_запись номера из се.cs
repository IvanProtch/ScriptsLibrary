//Заполнение номера операции в сквозном МО
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Linq;
using Intermech;
using Intermech.Expert;
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
        public int WorkTypeIndex { get; set; }

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
            No = Convert.ToInt32(row["cad009e6-306c-11d8-b4e9-00304f19f545" /*Номер объекта*/]);
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
        //#region test


        ////тест
        //List<NormForOperation> tpnorm1Ops = new List<NormForOperation>()
        //{
        //    new NormForOperation() { Caption = "тпнорм1оп1", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм1оп2", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм1оп3", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм1оп4", PartTime = "1 мин"},
        //};
        //List<Operation> tp1Ops = new List<Operation>()
        //{
        //    new Operation() { Caption = "тп1оп1", No = 5, Childs = new List<Item>(){tpnorm1Ops[0]}},
        //    new Operation() { Caption = "тп1оп2", No = 10, Childs = new List<Item>(){tpnorm1Ops[1]}},
        //    new Operation() { Caption = "тп1оп3", No = 15, Childs = new List<Item>(){tpnorm1Ops[2]}},
        //    new Operation() { Caption = "тп1оп4", No = 20, Childs = new List<Item>(){tpnorm1Ops[3]}},
        //};
        //tpnorm1Ops[0].Parents = new List<Item>() { tp1Ops[0] };
        //tpnorm1Ops[1].Parents = new List<Item>() { tp1Ops[1] };
        //tpnorm1Ops[2].Parents = new List<Item>() { tp1Ops[2] };
        //tpnorm1Ops[3].Parents = new List<Item>() { tp1Ops[3] };

        //TechProcess tp1_eco = new TechProcess() { Caption = "тп1", WorkTypeDescription = "88"};
        //List<NormForOperation> tpnorm2Ops = new List<NormForOperation>()
        //{
        //    new NormForOperation() { Caption = "тпнорм2оп1", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм2оп2", PartTime = "0 мин"},
        //    new NormForOperation() { Caption = "тпнорм2оп3", PartTime = ""},
        //};
        //List<Operation> tp2Ops = new List<Operation>()
        //{
        //    new Operation() { Caption = "тп2оп1", No = 5, Childs = new List<Item>(){tpnorm2Ops[0]} },
        //    new Operation() { Caption = "тп2оп2", No = 10, Childs = new List<Item>(){tpnorm2Ops[1]} },
        //    new Operation() { Caption = "тп2оп3", No = 15, Childs = new List<Item>(){tpnorm2Ops[2]} },
        //};
        //tpnorm2Ops[0].Parents = new List<Item>() { tp2Ops[0] };
        //tpnorm2Ops[1].Parents = new List<Item>() { tp2Ops[1] };
        //tpnorm2Ops[2].Parents = new List<Item>() { tp2Ops[2] };

        //TechProcess tp2_eco = new TechProcess() { Caption = "тп2", WorkTypeDescription = "77"};

        ////сборка
        //Item asm = new Item() { Caption = "asm", Parents = new List<Item>(), ObjectType = 1074};
        //Item mo1 = new Item() { Caption = "mo1", Parents = new List<Item>(){asm}, ObjectType = MetaDataHelper.GetObjectTypeID(SystemGUIDs.objtypeRouteProcessing) };
        //Item mo2 = new Item() { Caption = "mo2", Parents = new List<Item>(){asm}, ObjectType = MetaDataHelper.GetObjectTypeID(SystemGUIDs.objtypeRouteProcessing)  };
        //asm.Childs = new List<Item>() { mo1 };

        //List<NormForOperation> tpnorm3Ops = new List<NormForOperation>()
        //{
        //    new NormForOperation() { Caption = "тпнорм3оп1", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм3оп2", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм3оп3", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм3оп4", PartTime = "1 мин"},
        //};
        //List<Operation> tp3Ops = new List<Operation>()
        //{
        //    new Operation() { Caption = "тп1оп1", No = 5, Childs = new List<Item>(){tpnorm3Ops[0]}, OpNoInMO = "88.1.005"},
        //    new Operation() { Caption = "тп1оп2", No = 10, Childs = new List<Item>(){tpnorm3Ops[1]}, OpNoInMO = "88.1.010"},
        //    new Operation() { Caption = "тп1оп3", No = 15, Childs = new List<Item>(){tpnorm3Ops[2]}, OpNoInMO = "88.1.015"},
        //    new Operation() { Caption = "тп1оп4", No = 20, Childs = new List<Item>(){tpnorm3Ops[3]}, OpNoInMO = "88.1.020"},
        //};
        //tpnorm3Ops[0].Parents = new List<Item>() { tp3Ops[0] };
        //tpnorm3Ops[1].Parents = new List<Item>() { tp3Ops[1] };
        //tpnorm3Ops[2].Parents = new List<Item>() { tp3Ops[2] };
        //tpnorm3Ops[3].Parents = new List<Item>() { tp3Ops[3] };

        //TechProcess tp3 = new TechProcess() { Caption = "тп3", WorkTypeDescription = "88" };
        //List<NormForOperation> tpnorm4Ops = new List<NormForOperation>()
        //{
        //    new NormForOperation() { Caption = "тпнорм4оп1", PartTime = "1 мин"},
        //    new NormForOperation() { Caption = "тпнорм4оп2", PartTime = "0 мин"},
        //    new NormForOperation() { Caption = "тпнорм4оп3", PartTime = ""},
        //};
        //List<Operation> tp4Ops = new List<Operation>()
        //{
        //    new Operation() { Caption = "тп4оп1", No = 5, Childs = new List<Item>(){tpnorm4Ops[0]}, OpNoInMO = "77.1.005"},
        //    new Operation() { Caption = "тп4оп2", No = 10, Childs = new List<Item>(){tpnorm4Ops[1]}, OpNoInMO = "77.1.005" },
        //    new Operation() { Caption = "тп4оп3", No = 15, Childs = new List<Item>(){tpnorm4Ops[2]} , OpNoInMO = "77.1.005"},
        //};
        //tpnorm4Ops[0].Parents = new List<Item>() { tp4Ops[0] };
        //tpnorm4Ops[1].Parents = new List<Item>() { tp4Ops[1] };
        //tpnorm4Ops[2].Parents = new List<Item>() { tp4Ops[2] };

        //TechProcess tp4 = new TechProcess() { Caption = "тп4", WorkTypeDescription = "77" };

        //Item workShop1 = new Item() { Caption = "Цех1", Childs = tp1Ops.Cast<Item>().ToList(), Parents = new List<Item>() { tp1_eco } };
        //Item workShop2 = new Item() { Caption = "Цех2", Childs = tp2Ops.Cast<Item>().ToList(), Parents = new List<Item>() { tp2_eco } };
        //Item workShop3 = new Item() { Caption = "Цех3", Childs = tp3Ops.Cast<Item>().ToList(), Parents = new List<Item>() { tp3 } };
        //Item workShop4 = new Item() { Caption = "Цех4", Childs = tp4Ops.Cast<Item>().ToList(), Parents = new List<Item>() { tp4 } };
        //tp1_eco.Childs = new List<Item>() { workShop1 };
        //tp2_eco.Childs = new List<Item>() { workShop2 };
        //tp3.Childs = new List<Item>() { workShop3 };
        //tp4.Childs = new List<Item>() { workShop4 };


        //mo1.Childs = new List<Item>() {tp1_eco, tp2_eco, tp3, tp4};
        //tp1_eco.Parents = new List<Item>() { mo1 };
        //tp2_eco.Parents = new List<Item>() { mo1 };
        //tp3.Parents = new List<Item>() { mo1 };
        //tp4.Parents = new List<Item>() { mo1 };

        //#endregion

        if (System.Diagnostics.Debugger.IsAttached)
            System.Diagnostics.Debugger.Break();

        string error = string.Empty;

        foreach (var attachment in activity.Attachments)
        {
            //тп с непронумерованными операциями из вложения
            Dictionary<long, Item> ecoConsistance = new Dictionary<long, Item>();

            LoadComposition(activity.Session, new List<long>() { attachment.ObjectID }, ecoConsistance);

            //состав сборок, в которые входят операции из ии
            var ecoTPsApplicabilityASMsConsistance = LoadAsmsComposition(activity.Session, ecoConsistance
                .Values.OfType<TechProcess>()
                .Select(e => e.ObjectId).ToList());

            var ecoTechProcesses = ecoConsistance.Values.OfType<TechProcess>();

            ////тестовый вариант:
            //ecoTechProcesses = new List<TechProcess>() { tp1_eco, tp2_eco };
            //var ecoTPsApplicabilityASMsConsistance1 = new Dictionary<string, Item>();
            //ecoTPsApplicabilityASMsConsistance1[tp1_eco.Caption] = tp1_eco;
            //ecoTPsApplicabilityASMsConsistance1[tp2_eco.Caption] = tp2_eco;
            //ecoTPsApplicabilityASMsConsistance1[tp3.Caption] = tp3;
            //ecoTPsApplicabilityASMsConsistance1[tp4.Caption] = tp4;


            foreach (var ecoTP in ecoTechProcesses.OrderBy(e => e.Caption))
            {
                //находим соответствующий техпроцесс из ии в составе сборки
                if (ecoTPsApplicabilityASMsConsistance.ContainsKey(ecoTP.ObjectId))
                {
                    foreach (var productObj in ecoTPsApplicabilityASMsConsistance[ecoTP.ObjectId].EntersIn)
                    {
                        Dictionary<string, int> workTypeDescCounts = new Dictionary<string, int>();//вид производства, количество тп с этим видом

                        foreach (var mo in productObj.Value.Childs.Where(e => e.ObjectType == MetaDataHelper.GetObjectTypeID(SystemGUIDs.objtypeRouteProcessing)))
                        {
                            foreach (var tp in mo.Childs.Cast<TechProcess>())
                            {
                                //такой же тп, как в извещении, пропускаем, чтобы не влиял на общую нумерацию в таблице
                                if (tp.ObjectId == ecoTP.ObjectId)
                                    continue;

                                if (tp.Childs.FirstOrDefault() == null)
                                    continue;
                                if (tp.Childs.First().Childs.FirstOrDefault() == null)
                                    continue;

                                Operation firstOperation = tp.Childs.First().Childs.First() as Operation;

                                if (firstOperation.OpNoInMO == null)
                                    continue;

                                if (firstOperation.OpNoInMO.Length > 0)
                                {
                                    //записываем количество тп с одинаковым видом производства
                                    if (workTypeDescCounts.ContainsKey(tp.WorkTypeDescription))
                                        workTypeDescCounts[tp.WorkTypeDescription]++;
                                    else
                                        workTypeDescCounts[tp.WorkTypeDescription] = 1;
                                }
                            }
                        }
                        if (workTypeDescCounts.ContainsKey(ecoTP.WorkTypeDescription))
                            ecoTP.WorkTypeIndex = workTypeDescCounts[ecoTP.WorkTypeDescription] + 1;
                        else
                            ecoTP.WorkTypeIndex = 1;
                    }
                }
            }
            foreach (TechProcess techprocess in ecoConsistance.Values.OfType<TechProcess>())
                techprocess.CalculateOperationsNoInMO();

            foreach (Operation operation in ecoConsistance.Values.OfType<Operation>())
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

    /// <summary>
    /// Находим применяемость всех объектов до СЕ по связи тех. состав, у СЕ берем тех. состав; записываем в виде иерархии
    /// </summary>
    /// <param name="session"></param>
    /// <param name="objID"></param>
    /// <returns></returns>
    private Dictionary<long, Item> LoadAsmsComposition(IUserSession session, List<long> ids)
    {
        Dictionary<long, Item> result = new Dictionary<long, Item>();

        List<int> relIDs = new List<int>()
        {
            MetaDataHelper.GetRelationTypeID(new Guid("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/)),
        };

        List<int> childObjIDs = new List<int>()
        {
            MetaDataHelper.GetObjectTypeID("cad00132-306c-11d8-b4e9-00304f19f545" /*Сборочные единицы*/)
        };

        childObjIDs.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)));

        List<ColumnDescriptor> columns = new List<ColumnDescriptor>
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.CAPTION, AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_TYPE, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_PROJ_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.FieldName, SortOrders.NONE, 0),
        };

        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;

        DataTable dt = _loadService.LoadComplexCompositions(
        session.SessionGUID,
        ids.Select(e => new ObjInfoItem(e)).ToArray(),
        relIDs,
        childObjIDs,
        columns,
        false, // Режим получения данных (true - состав, false - применяемость)
        false, // Флаг группировки данных
        null, // Правило подбора версий
        null, // Условия на объекты при получении состава / применяемости
              //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
              //SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
        Intermech.SystemGUIDs.filtrationLatestVersions,
        null,
        -1 // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
        );

        foreach (DataRow row in dt.Rows)
        {
            new Item(row, result);
        }

        LoadComposition(session, dt.Rows.OfType<DataRow>().Select(e => (long)e["F_OBJECT_ID"]).ToList(), result);

        return result;
    }

}