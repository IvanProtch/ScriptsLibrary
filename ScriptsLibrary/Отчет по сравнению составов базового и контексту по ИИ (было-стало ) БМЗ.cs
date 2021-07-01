///Скрипт v.2.1 от 28.09.2020 Убраны все static для перехода на IPS 6

using System;
using System.Linq;
using System.Xml;
using Intermech.Interfaces;
using Intermech.Interfaces.Document;
using Intermech.Expert;
using Intermech.Kernel.Search;
using System.Data;
using Intermech.Document.Model;
using System.Collections.Generic;
using Intermech.Interfaces.Compositions;
using System.Collections.Specialized;
using System.Collections;
using System.Drawing;
using Intermech.Expert.Scenarios;
using Intermech.Interfaces.Client;
using Intermech.Navigator.Interfaces.QuickSearch;
using Intermech.Interfaces.Contexts;
using Intermech.Collections;
using Intermech;

namespace EcoDiffReport
{



    public class Script
    {
        
        public ICSharpScriptContext ScriptContext { get; private set; }
        public ScriptResult Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            Script1 sc = new Script1();
            sc.Execute(session, document, objectIDs);
            return new ScriptResult(true, document); 
        }
    }

    public class Script1
    {
        public bool Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            //try
            {
                bool compliteReport = true;
                document.Designation = "Сравнение составов";
                if (System.Diagnostics.Debugger.IsAttached)
                    System.Diagnostics.Debugger.Break();
                string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
                if (System.IO.File.Exists(file))
                    System.IO.File.Delete(file);
                AddToLog("Запускаем скрипт v 2.3");

                standart = MetaDataHelper.GetObjectTypeID("cad00252-306c-11d8-b4e9-00304f19f545"); //Стандартные

                proch = MetaDataHelper.GetObjectTypeID("cad0038d-306c-11d8-b4e9-00304f19f545"); //Прочие

                zagot = MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545"); //Заготовка

                matbase = MetaDataHelper.GetObjectTypeID("cad00170-306c-11d8-b4e9-00304f19f545"); //Материал базовый

                izdelie = MetaDataHelper.GetObjectTypeID("cad00268-306c-11d8-b4e9-00304f19f545" /*Изделия*/);

                sostMaterial = MetaDataHelper.GetObjectTypeID("cad00173-306c-11d8-b4e9-00304f19f545"); //Составной материал

                complects = MetaDataHelper.GetObjectTypeID("cad0025f-306c-11d8-b4e9-00304f19f545"); //Комплекты

                //Типы объектов на которые необходимо проверять заполнено ли количество
                checkKolvoTypes.Add(standart);
                checkKolvoTypes.Add(proch);
                checkKolvoTypes.Add(izdelie);
                checkKolvoTypes.Add(matbase);
                checkKolvoTypes.Add(sostMaterial);

                string ecoObjCaption = String.Empty;
                IDBObject ecoObj = session.GetObject(session.EditingContextID, false);
                if (ecoObj != null)
                {
                    var attr = ecoObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
                    if (attr != null)
                        document.Designation = attr.AsString;

                    ecoObjCaption = ecoObj.Caption;
                }

                document.Designation = document.Designation + "  " + DateTime.Now.ToString("dd.MM.yyyy, HH:mm:ss");
                document.DocumentName = document.Designation;
                //сборка по извещению
                IDBObject headerObj = null;
                //базовая версия
                IDBObject headerObjBase = null;
                var list = session.GetObjectIDVersions(objectIDs[0], false);
                foreach (var item in list)
                {
                    IDBObject obj = session.GetObject(item, true);
                    if (obj.ModificationID == session.EditingContextID)
                    {
                        headerObj = obj;
                    }

                    if (obj.IsBaseVersion)
                        headerObjBase = obj;
                }

                string headerObjCaption = String.Empty;
                if (headerObj == null)
                {
                    //Если нет созданной по извещению берем ту что подали на вход
                    headerObj = session.GetObject(objectIDs[0], true);
                    headerObjBase = headerObj;

                    headerObjCaption = headerObj.Caption;
                }

                if (headerObj == null)
                    throw new Exception("Не найдена версия для данного контекста редактирования");
                if (headerObjBase == null)
                    throw new Exception("Не найдена базовая версия");
                bool addBmzFields = true;
                
                #region Инициализация запроса к БД
                List<int> rels = new List<int>();
                rels.Add(MetaDataHelper.GetRelationTypeID("cad00023-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));
                rels.Add(MetaDataHelper.GetRelationTypeID("cad0019f-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));
                rels.Add(MetaDataHelper.GetRelationTypeID(
                "cad00584-306c-11d8-b4e9-00304f19f545" /*Состав экземпляров и партий изделий*/));

                List<ColumnDescriptor> columns = new List<ColumnDescriptor>();
                int attrId = (Int32)ObligatoryObjectAttributes.CAPTION;
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.FieldName, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));

                if (addBmzFields)
                {
                    attrId = MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
                }

                attrId = MetaDataHelper.GetAttributeTypeID("cad0038c-306c-11d8-b4e9-00304f19f545" /*Материал*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID("cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID("cad00020-306c-11d8-b4e9-00304f19f545" /*Наименование*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                if (addBmzFields)
                {
                    attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));

                    //attrId = MetaDataHelper.GetAttributeTypeID( "b1e25726-587e-4bb9-8934-75b85b57d5eb" /*Ответственный за номенклатуру ДМТО (БМЗ)*/);
                    //columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    //ColumnNameMapping.Guid, SortOrders.NONE, 0));

                    attrId = MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
                }

                attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstituteInGroup) /*Номер заменителя в группе*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));

                attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeManufacturingSign) /*Признак изготовления*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));

                //attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
                //List<ConditionStructure> conditions = new List<ConditionStructure>();
                //conditions.Add(new ConditionStructure(attrId, RelationalOperators.Equal, org_BMZ, LogicalOperators.NONE, 0, false));

                //columns.Add(new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/), AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.ID, SortOrders.NONE, 0));

                HybridDictionary tags = new HybridDictionary();
                DBRecordSetParams dbrsp = new DBRecordSetParams(null,
                columns != null
                ? columns.ToArray()
                : null,
                0, null,
                QueryConsts.All);
                dbrsp.Tags = tags;

                List<ObjInfoItem> items = new List<ObjInfoItem>();
                items.Add(new ObjInfoItem(headerObj.ObjectID));
                #endregion

                IDBEditingContextsService idbECS = session.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;
                var contextObject = idbECS.GetEditingContextsObject(session.SessionGUID, session.EditingContextID, true, false);
                if (contextObject != null)
                {
                    contextObjects = contextObject.GetVersionsID(session.EditingContextID, -1);
                    foreach (var objectid in contextObjects)
                    {
                        var info = session.GetObjectInfo(objectid);
                        contextObjectsID.Add(info.ID);
                    }
                }

                #region Первый состав по извещению
                DataTable dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
                Intermech.SystemGUIDs.filtrationBaseVersions, null);

                // Храним пару ид версии объекта + ид. физической величины
                // те объекты у которых посчитали количество
                // состав по извещен
                Dictionary<Tuple<long, long>, Item> ecoComposition = new Dictionary<Tuple<long, long>, Item>();

                AddToLog("Первый состав с извещением " + headerObj.ObjectID.ToString());
                if (dt != null)
                {
                    //dt.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script1.log");
                    //res2 это перезапсь itemsDict, в первый элемент кортежа пишется 0, во второй objectid. 
                    Dictionary<long, Item> itemsDict = new Dictionary<long, Item>();

                    Item header = new Item();
                    header.Id = headerObj.ID;
                    header.ObjectId = headerObj.ObjectID;
                    header.Caption = headerObj.Caption;
                    header.ObjectType = headerObj.ObjectType;
                    itemsDict[header.ObjectId] = header;

                    foreach (DataRow row in dt.Rows)
                    {
                        if (Convert.ToInt32(row["F_OBJECT_TYPE"]) == zagot &&
                            row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/].ToString() != "БМЗ")
                        {
                            continue;
                        }
                        var item = GetItem(row, itemsDict, true);
                        if (item != null)
                            AddToLog("item " + item.ToString());
                    }

                    itemsDict.Remove(headerObj.ObjectID);

                    //перезапись скопированных item в Composition
                    //откидывает item.MaterialId != 0 (те что не выполняют условия, другие типы) ???
                    foreach (var item in itemsDict.Values)
                    {
                        if (item.MaterialId != 0)
                        {
                            bool hasContextObjects = false;
                            var itemsKolvo = item.GetKolvo(true, ref hasContextObjects, ref item.HasEmptyKolvo);

                            Item mainClone = item.Clone();

                            if (!hasContextObjects)
                                continue;
                            if (itemsKolvo == null || itemsKolvo.Count == 0)
                            {
                                var kolvoItemClone = mainClone.Clone();

                                Item cachedItem;
                                var itemKey = new Tuple<long, long>(0, item.GetKey());
                                if (ecoComposition.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    ecoComposition[itemKey] = kolvoItemClone;
                                }

                                //res2[new Tuple<long, long>(0, item.GetKey())] = mainClone;

                                AddToLog("res2 " + mainClone);
                                continue;
                            }

                            foreach (var itemKolvo in itemsKolvo)
                            {
                                var kolvoItemClone = item.Clone();
                                kolvoItemClone.HasEmptyKolvo = mainClone.HasEmptyKolvo;
                                kolvoItemClone.KolvoSum = itemKolvo.Value;

                                Item cachedItem;
                                var itemKey = new Tuple<long, long>(itemKolvo.Key, item.GetKey());
                                if (ecoComposition.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    ecoComposition[itemKey] = kolvoItemClone;
                                }

                                //res2[item.ObjectId].KolvoSum = item.GetKolvo();
                                AddToLog("res2 " + kolvoItemClone);

                            }
                        }
                    }

                }
                else
                {
                    AddToLog("Состав1 не найден " + headerObj.ObjectID.ToString());
                }
                #endregion

                //сохраняем контекст редактирования
                long sessionid = session.EditingContextID;

                #region Второй состав без извещения
                AddToLog("Второй состав без извещения" + headerObjBase.ObjectID.ToString());
                session.EditingContextID = 0;

                items = new List<ObjInfoItem>();
                items.Add(new ObjInfoItem(headerObjBase.ObjectID));

                dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
                Intermech.SystemGUIDs.filtrationBaseVersions, null);

                Dictionary<Tuple<long, long>, Item> baseComposition = new Dictionary<Tuple<long, long>, Item>();
                if (dt != null)
                {
                    //dt.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script2.log");
                    Dictionary<long, Item> itemsDict = new Dictionary<long, Item>();
                    Item header = new Item();
                    header.Id = headerObjBase.ID;
                    header.ObjectId = headerObjBase.ObjectID;
                    header.Caption = headerObjBase.Caption;
                    header.ObjectType = headerObjBase.ObjectType;
                    itemsDict[header.ObjectId] = header;

                    foreach (DataRow row in dt.Rows)
                    {
                        if (Convert.ToInt32(row["F_OBJECT_TYPE"]) == zagot &&
                            row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/].ToString() != "БМЗ")
                        {
                            continue;
                        }
                        var item = GetItem(row, itemsDict, false);
                        if (item != null)
                            AddToLog("item " + item.ToString());
                    }

                    itemsDict.Remove(headerObjBase.ObjectID);
                    foreach (var item in itemsDict.Values)
                    {
                        if (item.MaterialId != 0)
                        {
                            bool hasContextObjects = false;
                            var itemsKolvo = item.GetKolvo(true, ref hasContextObjects, ref item.HasEmptyKolvo);
                            Item mainClone = item.Clone();

                            if (!hasContextObjects)
                                continue;


                            if (itemsKolvo == null || itemsKolvo.Count == 0)
                            {
                                var kolvoItemClone = mainClone.Clone();

                                Item cachedItem;
                                var itemKey = new Tuple<long, long>(0, item.GetKey());
                                if (baseComposition.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    baseComposition[itemKey] = kolvoItemClone;
                                }

                                //res1[new Tuple<long, long>(0, item.GetKey())] = mainClone;

                                AddToLog("res1 " + mainClone);
                                continue;
                            }

                            foreach (var itemKolvo in itemsKolvo)
                            {
                                var kolvoItemClone = item.Clone();
                                kolvoItemClone.HasEmptyKolvo = mainClone.HasEmptyKolvo;
                                kolvoItemClone.KolvoSum = itemKolvo.Value;

                                Item cachedItem;
                                var itemKey = new Tuple<long, long>(itemKolvo.Key, item.GetKey());
                                if (baseComposition.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    baseComposition[itemKey] = kolvoItemClone;
                                }

                                //res1[item.ObjectId].KolvoSum = item.GetKolvo();
                                AddToLog("res1 " + kolvoItemClone);

                            }
                        }
                    }
                }
                else
                {
                    AddToLog("Состав2 не найден " + headerObjBase.ObjectID.ToString());
                }
                #endregion

                //возвращаем контекст
                session.EditingContextID = sessionid;

                List<Material> resultComposition = new List<Material>();

                foreach (var resItem in baseComposition)
                {
                    var baseItem = resItem.Value;
                    Material mat = new Material();
                    if (baseItem.MaterialDMTO != null)
                        mat.MaterialDMTO = baseItem.MaterialDMTO;
                    mat.MaterialId = baseItem.MaterialId;
                    mat.MaterialCode = baseItem.MaterialCode;
                    mat.MaterialCaption = baseItem.MaterialCaption;
                    AddToLog("item0 " + baseItem.ToString());
                    //добавляем базовый состав в результат
                    resultComposition.Add(mat);

                    mat.HasEmptyKolvo1 = baseItem.HasEmptyKolvo;
                    AddToLog("mat01 " + mat.ToString());
                    if (baseItem.KolvoSum != null)
                    {
                        AddToLog("item.KolvoSum != null");

                        //записываем количество из базовой версии
                        mat.Kolvo1 = baseItem.KolvoSum.Value;
                        mat.MeasureId = baseItem.KolvoSum.MeasureID;
                        var descr = MeasureHelper.FindDescriptor(mat.MeasureId);

                        foreach (var kolvoinasm1 in baseItem.KolvoInAsm)
                        {
                            var asmKolvo = kolvoinasm1.Key.Parent.ParentItems.FirstOrDefault() != null ? kolvoinasm1.Key.Parent.ParentItems.First().Kolvo : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                            mat.EntersInAsm1[kolvoinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(kolvoinasm1.Value, asmKolvo);
                        }

                        if (descr != null)
                            mat.EdIzm = descr.ShortName;
                        else
                        {
                            AddToLog("descr = null  item.KolvoSum = " + baseItem.KolvoSum.Caption + "  MeasureId = " +
                            mat.MeasureId);
                        }
                    }
                    else
                    {
                        AddToLog("item.KolvoSum == null");
                        mat.HasEmptyKolvo1 = true;
                    }

                    Item ecoItem;
                    //Находим базовый объект в списке объектов из извещения
                    if (ecoComposition.TryGetValue(resItem.Key, out ecoItem))
                    {
                        ecoComposition.Remove(resItem.Key);
                    }
                    else
                    {
                        var emptyItem = new Tuple<long, long>(0, resItem.Key.Item2);
                        if (ecoComposition.TryGetValue(emptyItem, out ecoItem))
                        {
                            ecoComposition.Remove(emptyItem);
                        }
                    }

                    // Отдельно обработаем случай с заготовками 
                    if (ecoItem == null
                    && MetaDataHelper.IsObjectTypeChildOf(resItem.Value.ObjectType, zagot))
                    {
                        Tuple<long, long> zagotItem = null;

                        foreach (var res2Item in ecoComposition)
                        {
                            if (res2Item.Key.Item1 != resItem.Key.Item1 ||
                            res2Item.Value.MaterialId != resItem.Value.MaterialId ||
                            !MetaDataHelper.IsObjectTypeChildOf(resItem.Value.ObjectType, zagot))
                            {
                                continue;
                            }

                            ecoItem = res2Item.Value;
                            zagotItem = res2Item.Key;
                        }

                        if (zagotItem != null)
                        {
                            ecoComposition.Remove(zagotItem);
                        }
                    }

                    if (ecoItem != null)
                    {
                        //mat.Kolvo2 = MeasureHelper.ConvertToMeasuredValue(item2.KolvoSum, mat.MeasureId).Value;
                        if (ecoItem.MaterialDMTO != null)
                            mat.MaterialDMTO = ecoItem.MaterialDMTO;
                        mat.HasEmptyKolvo2 = ecoItem.HasEmptyKolvo;
                        AddToLog("item02 " + ecoItem.ToString());
                        if (ecoItem.KolvoSum != null)
                        {
                            AddToLog("item2.KolvoSum != null");
                            mat.Kolvo2 = MeasureHelper.ConvertToMeasuredValue(ecoItem.KolvoSum, mat.MeasureId).Value;

                            foreach (var kolvoinasm1 in ecoItem.KolvoInAsm)
                            {
                                var asmKolvo = kolvoinasm1.Key.Parent.ParentItems.FirstOrDefault() != null ? kolvoinasm1.Key.Parent.ParentItems.First().Kolvo : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                                mat.EntersInAsm2[kolvoinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(kolvoinasm1.Value, asmKolvo);
                            }
                        }
                        else
                        {
                            AddToLog("item2.KolvoSum == null");
                            mat.HasEmptyKolvo2 = true;
                        }

                    }

                    AddToLog("mat1 " + mat.ToString());
                }

                //Добавим те которых не было в первом наборе
                foreach (var item in ecoComposition.Values) 
                {
                    Material mat = new Material();
                    if (item.MaterialDMTO != null)
                        mat.MaterialDMTO = item.MaterialDMTO;
                    mat.MaterialDMTO = item.MaterialDMTO;
                    mat.MaterialId = item.MaterialId;
                    mat.MaterialCode = item.MaterialCode;
                    mat.MaterialCaption = item.MaterialCaption;
                    //Добавляем снова
                    resultComposition.Add(mat);

                    mat.HasEmptyKolvo2 = item.HasEmptyKolvo;
                    if (item.KolvoSum != null)
                    {
                        //Кажется, единственное, что будет отличаться и должно отличаться -- количество
                        //Добавляем количество из версии по ии
                        mat.Kolvo2 = item.KolvoSum.Value;
                        mat.MeasureId = item.KolvoSum.MeasureID;
                        var descr = MeasureHelper.FindDescriptor(mat.MeasureId);

                        foreach (var kolvoinasm1 in item.KolvoInAsm)
                        {
                            var asmKolvo = kolvoinasm1.Key.Parent.ParentItems.FirstOrDefault() != null ? kolvoinasm1.Key.Parent.ParentItems.First().Kolvo : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                            mat.EntersInAsm2[kolvoinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(kolvoinasm1.Value, asmKolvo);
                        }

                        if (descr != null)
                            mat.EdIzm = descr.ShortName;
                        else
                        {
                            AddToLog("descr = null  item.KolvoSum = " + item.KolvoSum.Caption + "  MeasureId = " +
                            mat.MeasureId);
                        }
                    }
                    else
                    {
                        mat.HasEmptyKolvo2 = true;
                    }

                    AddToLog("mat2 " + mat.ToString());
                }

                List<long> ids = new List<long>();
                foreach (var item in resultComposition)
                {
                    ids.Add(item.MaterialId);
                    if (string.IsNullOrEmpty(item.MaterialCode))
                    {

                    }
                }

                //Получаем объекты из resultComposition у которых MaterialId != 0 и добавляем к соответствующему Material значение кода АМТО
                if (ids.Count > 0)
                {
                    IDBObjectCollection col = session.GetObjectCollection(matbase);
                    List<ConditionStructure> conds = new List<ConditionStructure>();
                    conds.Add(new ConditionStructure((Int32)ObligatoryObjectAttributes.F_ID, RelationalOperators.In,
                    ids.ToArray(), LogicalOperators.NONE, 0, false));

                    columns = new List<ColumnDescriptor>();
                    columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_ID, AttributeSourceTypes.Auto,
                    ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0)
                    );

                    columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_OBJECT_ID,
                    AttributeSourceTypes.Object,
                    ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0)
                    );
                    columns.Add(new ColumnDescriptor(
                    MetaDataHelper.GetAttributeTypeID("cad00020-306c-11d8-b4e9-00304f19f545" /*Наименование*/),
                    AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));
                    columns.Add(new ColumnDescriptor(
                    MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/),
                    AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));
                    //columns.Add(new ColumnDescriptor(MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/), AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.ID, SortOrders.NONE, 0));
                    tags = new HybridDictionary();
                    dbrsp = new DBRecordSetParams(conds.ToArray(),
                    columns != null
                    ? columns.ToArray()
                    : null,
                    0, null,
                    QueryConsts.All);
                    dt = col.Select(dbrsp);
                    foreach (DataRow row in dt.Rows)
                    {
                        long id = Convert.ToInt64(row["F_ID"]);
                        string name = Convert.ToString(row["cad00020-306c-11d8-b4e9-00304f19f545"]);
                        string code = Convert.ToString(row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/]);
                        Material mat = resultComposition.Find(x => x.MaterialId == id);
                        if (mat != null)
                        {
                            mat.MaterialCode = code;
                            mat.MaterialCaption = name;
                            AddToLog("mat3 " + mat.ToString());
                        }
                    }

                    // там где код АМТО == ""
                    foreach (var item in resultComposition)
                    {
                        if (item.MaterialCode == "")
                        {
                            item.MaterialCode = "";
                            IDBObject obj = session.GetObject(item.MaterialId, false);
                            if (obj != null)
                            {
                                IMSObjectType type = MetaDataHelper.GetObjectType(obj.ObjectType);

                                IDBAttribute attr =
                                obj.GetAttributeByGuid(new Guid("cad00020-306c-11d8-b4e9-00304f19f545"/*Наименование*/), false);

                                //записываем наименование и код АМТО ?
                                if (attr != null)
                                    item.MaterialCaption = attr.AsString;
                                attr = obj.GetAttributeByGuid(new Guid("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/), false);
                                if (attr != null)
                                    item.MaterialCode = attr.AsString;
                            }

                            AddToLog("mat4 " + item.ToString());
                        }
                    }
                }

                foreach (var item in resultComposition)
                {

                    AddToLog("beforesort " + item.ToString());
                }

                AddToLog("res.Sort");

                resultComposition.Sort(Compare);

                foreach (var item in resultComposition)
                {

                    AddToLog("aftersort " + item.ToString());
                }

                List<Material> restemp = new List<Material>();
                foreach (var item in resultComposition)
                {
                    Material lastItem = null;
                    if (restemp.Count > 0)
                        lastItem = restemp[restemp.Count - 1];
                    if (!string.IsNullOrWhiteSpace(item.MaterialCode))
                    {
                        if (lastItem != null && lastItem.MaterialCode == item.MaterialCode && lastItem.EdIzm == item.EdIzm)
                        {
                            lastItem.Kolvo1 += item.Kolvo1;
                            lastItem.Kolvo2 += item.Kolvo2;
                            lastItem.HasEmptyKolvo1 |= item.HasEmptyKolvo1;
                            lastItem.HasEmptyKolvo2 |= item.HasEmptyKolvo2;
                        }
                        else
                        {
                            restemp.Add(item);
                        }
                    }
                    else
                    {
                        restemp.Add(item);
                    }
                }

                resultComposition = restemp;


                #region Формирование отчета
                
                // Заполнение шапки
                AddToLog("fill header");
                DocumentTreeNode headerNode = document.Template.FindNode("Шапка");
                if (headerNode != null)
                {
                    TextData textData = headerNode.FindNode("Извещение") as TextData;
                    if (textData != null)
                    {
                        textData.AssignText(ecoObjCaption, false, false, false);
                    }

                    textData = headerNode.FindNode("ДСЕ") as TextData;
                    if (textData != null)
                    {
                        textData.AssignText(headerObjCaption, false, false, false);
                    }
                }

                DocumentTreeNode docrow = document.Template.FindNode("Строка");
                DocumentTreeNode table = document.FindNode("Таблица");
                DocumentTreeNode doccol2 = document.Template.FindNode("Столбец2");
                DocumentTreeNode docrow2 = document.Template.FindNode("Строка2");


                int index = 0;
                int i = 0;
                foreach (var item in resultComposition)
                {
                    {
                        if (item.Kolvo1 != item.Kolvo2 || item.HasEmptyKolvo1 != item.HasEmptyKolvo2)
                        {

                            AddToLog("createnode " + item.ToString());
                            DocumentTreeNode node = docrow.CloneFromTemplate(true, true);
                            if (compliteReport)
                            {
                                //оставляем только различающиеся элементы EntersInAsm1 и EntersInAsm2
                                var keys1 = item.EntersInAsm1.Keys.ToList();
                                foreach (var key in keys1)
                                {
                                    Tuple<MeasuredValue, MeasuredValue> value = null;
                                    if (item.EntersInAsm2.TryGetValue(key, out value))
                                    {
                                        if (value != null)
                                        {
                                            if (value.Item1.Value == item.EntersInAsm1[key].Item1.Value)
                                            {
                                                item.EntersInAsm1.Remove(key);
                                                item.EntersInAsm2.Remove(key);
                                            }
                                        }
                                    }
                                }

                                #region Запись "было"
                                if(item.EntersInAsm1.Count > 0 || item.EntersInAsm2.Count > 0)
                                {
                                    index++;
                                    i++;
                                    table.AddChildNode(node, false, false);

                                    Write(node, "Индекс", index.ToString());
                                    Write(node, "Код", item.MaterialCode);
                                    Write(node, "Материал", item.MaterialCaption);

                                    if (item.Kolvo1 != 0 || !item.HasEmptyKolvo1)
                                    {
                                        Write(node, "Всего", Math.Round(item.Kolvo1, 4).ToString() + " " + item.EdIzm);


                                        if (index == 1)
                                        {
                                            //DocumentTreeNode entrow = node.FindNode("Куда входит строка");
                                            DocumentTreeNode row2 = node.FindNode("Строка2");
                                            if (item.EntersInAsm1.Count != 0)
                                            {
                                                WriteFirstAfterParent(row2, "Куда входит", item.EntersInAsm1.First().Key);
                                                WriteFirstAfterParent(row2, "Количество вхождений", item.EntersInAsm1.First().Value.Item1.Caption);
                                                WriteFirstAfterParent(row2, "Количество сборок", item.EntersInAsm1.First().Value.Item2.Caption);
                                            }

                                            for (int j = 1; j < item.EntersInAsm1.Count; j++)
                                            {
                                                DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                                row2.AddChildNode(node2, false, false);
                                                WriteFirstAfterParent(node2, "Куда входит", item.EntersInAsm1.Keys.ToList()[j]);
                                                WriteFirstAfterParent(node2, "Количество вхождений", item.EntersInAsm1.Values.ToList()[j].Item1.Caption);
                                                WriteFirstAfterParent(node2, "Количество сборок", item.EntersInAsm1.Values.ToList()[j].Item2.Caption);
                                            }
                                        }

                                        if (index > 1)
                                        {
                                            //DocumentTreeNode entrow = node.FindNode(string.Format("Куда входит строка #{0}", index));
                                            //DocumentTreeNode entcol = node.FindNode(string.Format("Куда входит #{0}", index));
                                            DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", i));
                                            if (item.EntersInAsm1.Count != 0)
                                            {
                                                WriteFirstAfterParent(row2, string.Format("Куда входит #{0}", i), item.EntersInAsm1.First().Key);
                                                WriteFirstAfterParent(row2, string.Format("Количество вхождений #{0}", i), item.EntersInAsm1.First().Value.Item1.Caption);
                                                WriteFirstAfterParent(row2, string.Format("Количество сборок #{0}", i), item.EntersInAsm1.First().Value.Item2.Caption);
                                            }

                                            for (int j = 1; j < item.EntersInAsm1.Count; j++)
                                            {

                                                DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                                DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", i));

                                                i++;
                                                col2.AddChildNode(node2, false, false);
                                                WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", i), item.EntersInAsm1.Keys.ToList()[j]);
                                                WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", i), item.EntersInAsm1.Values.ToList()[j].Item1.Caption);
                                                WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", i), item.EntersInAsm1.Values.ToList()[j].Item2.Caption);

                                            }

                                        }
                                    }
                                }
                                #endregion


                                #region Запись "стало"

                                if(item.EntersInAsm1.Count > 0 || item.EntersInAsm2.Count > 0)
                                {
                                    i++;
                                    node = docrow.CloneFromTemplate(true, true);

                                    table.AddChildNode(node, false, false);

                                    if (item.Kolvo2 != 0 || !item.HasEmptyKolvo2)
                                    {
                                        Write(node, "Всего", Math.Round(item.Kolvo1, 4).ToString() + " " + item.EdIzm);

                                        DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", i));
                                        DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", i));
                                        if (item.EntersInAsm2.Count != 0)
                                        {
                                            WriteFirstAfterParent(row2, string.Format("Куда входит #{0}", i), item.EntersInAsm2.First().Key);
                                            WriteFirstAfterParent(row2, string.Format("Количество вхождений #{0}", i), item.EntersInAsm2.First().Value.Item1.Caption);
                                            WriteFirstAfterParent(row2, string.Format("Количество сборок #{0}", i), item.EntersInAsm2.First().Value.Item2.Caption);
                                        }

                                        for (int j = 1; j < item.EntersInAsm2.Count; j++)
                                        {
                                            DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                            col2.AddChildNode(node2, false, false);
                                            i++;
                                            WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", i), item.EntersInAsm2.Keys.ToList()[j]);
                                            WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", i), item.EntersInAsm2.Values.ToList()[j].Item1.Caption);
                                            WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", i), item.EntersInAsm2.Values.ToList()[j].Item2.Caption);

                                        }
                                    }
                                }
                                
                                #endregion


                            }
                            else
                            {
                                index++;
                                table.AddChildNode(node, false, false);

                                Write(node, "Индекс", index.ToString());
                                Write(node, "ДМТО", item.MaterialDMTO);

                                Write(node, "Код", item.MaterialCode);
                                Write(node, "Материал", item.MaterialCaption);
                                if (item.Kolvo1 != 0 || !item.HasEmptyKolvo1)
                                    Write(node, "Было", Convert.ToString(Math.Round(item.Kolvo1, 4)));
                                else
                                    Write(node, "Было", "-");
                                if (item.Kolvo2 != 0 || !item.HasEmptyKolvo2)
                                    Write(node, "Будет", Convert.ToString(Math.Round(item.Kolvo2, 4)));
                                else
                                    Write(node, "Будет", "-");
                                var diff = Math.Round(item.Kolvo2 - item.Kolvo1, 4);
                                Write(node, "Разница", Convert.ToString(diff)/*Convert.ToString(item.Kolvo2 - item.Kolvo1)*/);
                                Write(node, "ЕдИзм", item.EdIzm);
                            }

                            if (item.HasEmptyKolvo1 || item.HasEmptyKolvo2)
                            {
                                (node as RectangleElement).AssignLeftBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                (node as RectangleElement).AssignRightBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                (node as RectangleElement).AssignTopBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                (node as RectangleElement).AssignBottomBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                AddToLog("item.HasEmptyKolvo  " + item.ToString());
                                foreach (DocumentTreeNode child in node.Nodes)
                                {
                                    (child as RectangleElement).AssignLeftBorderLine(
                                    new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                    (child as RectangleElement).AssignRightBorderLine(
                                    new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                    (child as RectangleElement).AssignTopBorderLine(
                                    new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                    (child as RectangleElement).AssignBottomBorderLine(
                                    new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                    if (child is TextData)
                                    {
                                        //AddToLog("child  id = " + child.Id);
                                        if ((child as TextData).CharFormat != null)
                                        {
                                            CharFormat cf = (child as TextData).CharFormat.Clone();
                                            cf.TextColor = Color.Red;
                                            cf.CharStyle = CharStyle.Bold;
                                            //AddToLog("SetCharFormat");
                                            (child as TextData).SetCharFormat(cf, false, false);
                                        }
                                        else
                                        {
                                            //AddToLog("(child as TextData).CharFormat == null");
                                        }
                                    }
                                }
                            }
                        }
                    }
                        
                }

                AddToLog("Завершили создание отчета ");
                #endregion
            }
            //catch (Exception ex)
            //{
            //AddToLog(ex.Message);
            //AddToLog(ex.StackTrace);
            //throw ex;
            //}
            return true;
        }

        int Compare(Material x, Material y)
        {
            int res = x.MaterialCode.CompareTo(y.MaterialCode);
            if (res == 0)
            {

                res = x.EdIzm.CompareTo(y.EdIzm);
            }
            return res;
        }

        //Список идентификаторов версий объектов
        public List<long> contextObjects = new List<long>();
        //Список идентификаторов  объектов
        public List<long> contextObjectsID = new List<long>();
        public List<int> checkKolvoTypes = new List<int>();

        public void AddToLog(string text)
        {
            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
            text = text + Environment.NewLine;
            System.IO.File.AppendAllText(file, text);
            //AddToOutputView(text);
        }

        //public static void AddToOutputView(string text)
        //{
        //    IOutputView view = ServicesManager.GetService(typeof(IOutputView)) as IOutputView;
        //    if (view != null)
        //        view.WriteString("Скрипт сравнения составов", text);
        //    else
        //        AddToLog("view == null");
        //}

        private DocumentTreeNode AddNode(DocumentTreeNode childNode, string value)
        {
            DocumentTreeNode node = childNode.GetDocTreeRoot().Template.FindNode(value).CloneFromTemplate(true, true);
            childNode.AddChildNode(node, false, false);
            return node;
        }

        private void Write(DocumentTreeNode parent, string tplid, string text)
        {
            TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
            if (td != null)
                td.AssignText(text, false, false, false);
        }

        private void WriteFirstAfterParent(DocumentTreeNode parent, string tplid, string text)
        {
            TextData td = parent.FindNode(tplid) as TextData;
            if (td != null)
                td.AssignText(text, false, false, false);
        }

        int izdelie;
        int standart;
        int complects;
        int proch;
        int zagot;
        int matbase;
        int sostMaterial;

        /// <summary>
        /// Если не выполняются определенные условия (н-р, не покупное изд), materialid=0, и дальше объект пропускается
        /// </summary>
        /// <param name="row"></param>
        /// <param name="itemsDict">Словарь в который будут заноситься objectid и Item. Организует иерархическую связь между Item.</param>
        /// <param name="contextMode"></param>
        /// <returns></returns>
        Item GetItem(DataRow row, Dictionary<long, Item> itemsDict, bool contextMode)
        {
            Item item = null;
            long id = Convert.ToInt64(row["F_OBJECT_ID"]);
            if (itemsDict.ContainsKey(id))
            {
                item = itemsDict[id];
            }
            else
            {

                item = new Item();
                item.Id = Convert.ToInt64(row["F_PART_ID"]);
                item.ObjectId = Convert.ToInt64(row["F_OBJECT_ID"]);
                item.ObjectType = Convert.ToInt32(row["F_OBJECT_TYPE"]);
                string designation = Convert.ToString(row[Intermech.SystemGUIDs.attributeDesignation]);
                string name = Convert.ToString(row[Intermech.SystemGUIDs.attributeName]);
                string caption = name;
                if (!string.IsNullOrEmpty(designation))
                    caption = designation + " " + name;
                item.Caption = caption;
                item.ObjectCaption = Convert.ToString(row["CAPTION"]);
                if (!contextMode)
                {
                    if (contextObjects.Contains(item.ObjectId) || contextObjects.Contains(-item.ObjectId)) return null; //Объект входит в состав извещения значит считаем что в составе без извещения его нет
                }
                if (contextObjectsID.Contains(item.Id))
                    item.isContextObject = true;


                //object dmto = row[ "b1e25726-587e-4bb9-8934-75b85b57d5eb" /*Ответственный за номенклатуру ДМТО (БМЗ)*/];
                //if (dmto  != DBNull.Value)
                //{
                //	item.MaterialDMTO = Convert.ToString(dmto);
                //}
            }
            
            long parentId = Convert.ToInt64(row["F_PROJ_ID"]);
            Relation lnk = new Relation();
            lnk.Child = item;
            lnk.LinkId = Convert.ToInt64(row["F_PRJLINK_ID"]);
            lnk.RelationTypeId = Convert.ToInt32(row["F_RELATION_TYPE"]);

            // если нашли родительский объект в словаре состава, добавляем связи
            if (itemsDict.ContainsKey(parentId))
            {
                Item parent = itemsDict[parentId];
                if (parent.ObjectType == sostMaterial || MetaDataHelper.IsObjectTypeChildOf(parent.ObjectType, sostMaterial))
                {
                    return null;
                }
                if (parent.isPocup)
                {
                    AddToLog("ParentIsPocup " + parentId + "  " + lnk.ToString());
                    return null;
                }

                parent.ChildItems.Add(lnk);
                lnk.Parent = parent;
                item.ParentItems.Add(lnk);
            }
            else
            {
                AddToLog("ParentNotFound " + parentId + "  " + lnk.ToString());
                return null;
            }

            bool isDopZamen = false;
            object dopZamenValue = row[Intermech.SystemGUIDs.attributeSubstituteInGroup];
            if (dopZamenValue != DBNull.Value)
            {
                isDopZamen = Convert.ToInt32(dopZamenValue) != 0;
            }

            bool isPocup = false;
            object pocupValue = row[Intermech.SystemGUIDs.attributeManufacturingSign];
            if (pocupValue != DBNull.Value)
            {
                isPocup = Convert.ToInt32(pocupValue) != 1;
            }

            item.isPocup = isPocup;

            //if (isDopZamen)//Допзамены и их состав не считаем
            //{
            //    AddToLog("isDopZamen " + lnk.ToString());
            //    return null;
            //}

            itemsDict[id] = item;

            // получаем значение количества и записываем в Link
            object kolvoValue = row["cad00267-306c-11d8-b4e9-00304f19f545"];
            if (kolvoValue is string)
            {
                lnk.Kolvo = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(kolvoValue), false);
            }

            //if (isDopZamen && lnk.Kolvo != null)
            //lnk.Kolvo.Value = 0;

            if (lnk.Kolvo == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
            MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545")) != null)
            {
                foreach (var checkKolvoType in checkKolvoTypes)
                {
                    if (lnk.Child != null && (lnk.Child.ObjectType == checkKolvoType ||
                    MetaDataHelper.IsObjectTypeChildOf(lnk.Child.ObjectType, checkKolvoType)))
                    {
                        lnk.HasEmptyKolvo = true;
                    }
                }
            }

            long materialId = item.Id;
            string codeMaterial = Convert.ToString(row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/]);

            if (item.isPocup)
            {
                item.MaterialId = materialId;
                item.MaterialCode = codeMaterial;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == proch || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, proch))
            {
                item.MaterialId = materialId;
                item.MaterialCode = codeMaterial;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == standart || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, standart))
            {
                item.MaterialId = materialId;
                item.MaterialCode = codeMaterial;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == complects || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, complects))
            {
                item.MaterialId = materialId;
                item.MaterialCode = codeMaterial;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == matbase || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, matbase))
            {
                item.MaterialId = materialId;
                item.MaterialCode = codeMaterial;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == zagot)
            {
                kolvoValue = row["cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/];
                if (kolvoValue is string)
                {
                    lnk.Kolvo = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(kolvoValue), false);
                }

                if (lnk.Kolvo == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
                MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/)) != null)
                {
                    lnk.HasEmptyKolvo = true;
                }

                object matId = row["cad005e3-306c-11d8-b4e9-00304f19f545"];
                if (matId != DBNull.Value)
                {
                    materialId = Convert.ToInt64(matId);
                    item.MaterialId = materialId;
                }
            }

            AddToLog("CreateLink " + lnk.ToString());
            return item;
        }

        class Material
        {
            public long MaterialId;
            public string MaterialCaption;
            string materialCode;
            public string MaterialDMTO;
            /// <summary>
            /// Код АМТО
            /// </summary>
            public string MaterialCode
            {
                get
                {
                    return materialCode;
                }
                set
                {
                    if (value == null)
                        materialCode = "";
                    else
                        materialCode = value;
                }
            }
            /// <summary>
            /// Количество из базовой версии
            /// </summary>
            public double Kolvo1 = 0;
            /// <summary>
            /// Количество из версии по извещению
            /// </summary>
            public double Kolvo2 = 0;
            public long MeasureId;
            string edIzm = "";
            public string EdIzm
            {
                get
                {
                    return edIzm;
                }
                set
                {
                    if (value == null)
                        edIzm = "";
                    else
                        edIzm = value;
                }
            }
            public bool HasEmptyKolvo1 = false;
            public bool HasEmptyKolvo2 = false;

            /// <summary>
            /// Вхождения СЕ (название, кол-во, кол-во се)
            /// </summary>
            public Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> EntersInAsm1 = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();
            public Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> EntersInAsm2 = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();

            public override string ToString()
            {
                return string.Format("MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; Kolvo1={3}; Kolvo2={4};",
                MaterialId, MaterialCode, MaterialCaption, Kolvo1, Kolvo2);
            }
        }

        /// <summary>
        /// Связь между объектами Item. Через связь записывается исходное значение количества материала
        /// </summary>
        class Relation
        {
            public long LinkId;
            public int RelationTypeId;
            public Item Child;
            public Item Parent;
            public MeasuredValue Kolvo;
            public bool HasEmptyKolvo = false;

            public IDictionary<long, MeasuredValue> GetKolvo(ref bool hasContextObject, ref bool hasemptyKolvoRelations)
            {
                IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
                IDictionary<long, MeasuredValue> itemsKolvo = Parent.GetKolvo(false, ref hasContextObject, ref hasemptyKolvoRelations);
                //Количество инициализируется в методе GetItem
                // если значение количества пустое, записываем количество у связи, затем возвращаем
                if (Kolvo != null)
                {
                    if (itemsKolvo == null || itemsKolvo.Count == 0)
                    {
                        result[MeasureHelper.FindDescriptor(Kolvo).PhysicalQuantityID] = Kolvo.Clone() as MeasuredValue;
                    }
                    else
                    {
                        try
                        {
                            foreach (var itemKolvo in itemsKolvo)
                            {
                                itemKolvo.Value.Multiply(Kolvo);
                                result[MeasureHelper.FindDescriptor(Kolvo).PhysicalQuantityID] = itemKolvo.Value;
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(this.ToString() + " " + ex.Message, ex);
                        }
                    }
                }
                else
                {

                    if (HasEmptyKolvo)
                    {
                        hasemptyKolvoRelations = true;
                        if (Parent == null)
                        {
                            AddToLogForLink("Parent == null" + LinkId);
                        }

                        if (Child == null)
                        {
                            AddToLogForLink("Child == null" + LinkId);
                        }

                        if (Parent != null && Child != null)
                        {
                            string text = "Отсутствует количество. Тип связи " +
                            MetaDataHelper.GetRelationTypeName(RelationTypeId) + " Род. объект " +
                            Parent.ToString() + " Дочерний объект " + Child.ToString();
                            string text1 = "Отсутствует количество. " + " Род. объект '" + Parent.ObjectCaption +
                            "' Дочерний объект '" + Child.ObjectCaption + "'";
                            
                            //Script1.AddToOutputView(text1);
                            #region Перенесем выполнение кода статического метода в текущее место
                            IOutputView view = ServicesManager.GetService(typeof(IOutputView)) as IOutputView;
                            if (view != null)
                                view.WriteString("Скрипт сравнения составов", text1);
                            else
                                AddToLogForLink("view == null");

                            #endregion
                            AddToLogForLink(text);
                        }
                    }

                    if (Child.ChildItems == null || Child.ChildItems.Count == 0)
                    {
                        //Если это материал и он последний, то его количество равно 0
                        return result; // null
                    }

                    if (itemsKolvo != null)
                    {
                        result = itemsKolvo;
                    }
                }

                return result;
            }
            public void AddToLogForLink(string text)
            {
                string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
                text = text + Environment.NewLine;
                System.IO.File.AppendAllText(file, text);
                //AddToOutputView(text);
            }


            public override string ToString()
            {
                string res = "Link=" + LinkId.ToString();
                if (Parent != null)
                    res = res + " parent= " + Parent.ToString();
                if (Child != null)
                    res = res + " child= " + Child.ToString();
                res = res + " Количество = " + Convert.ToString(Kolvo);
                return res;
            }
        }

        /// <summary>
        /// Объект. Хранит ссылки на связи Link; сумму количества материала (получено из связей)
        /// </summary>
        class Item
        {
            public override string ToString()
            {
                string objType1 = MetaDataHelper.GetObjectTypeName(ObjectType);
                //return string.Format( " Id={3}; Caption={4}; MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; KolvoSum={5}; ObjectType = {6}", MaterialId, MaterialCode, MaterialCaption, Id, Caption, KolvoSum, objType1);
                return string.Format(
                " Id={3}; Caption={4}; MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; KolvoSum={5}; ObjectType = {6}",
                MaterialId, MaterialCode, MaterialCaption, Id, Caption, KolvoSum, objType1);
            }
            
            public Item Clone()
            {
                Item clone = new Item();
                clone.MaterialId = MaterialId;
                clone.MaterialCaption = MaterialCaption;
                clone.MaterialCode = MaterialCode;
                clone.MaterialDMTO = MaterialDMTO;
                clone.Id = Id;
                clone.ObjectId = ObjectId;
                clone.Caption = Caption;
                clone.ObjectGuid = ObjectGuid;
                clone.ObjectType = ObjectType;
                clone.ParentItems = ParentItems;
                clone.ChildItems = ChildItems;
                clone._kolvoInAsm = _kolvoInAsm;

                //     clone.Kolvo = Kolvo;
                return clone;
            }

            public string MaterialDMTO = null;
            public bool isContextObject = false;
            public bool isPocup = false;
            public MeasuredValue KolvoSum;
            /// <summary>
            /// MaterialId == Id ?
            /// </summary>
            public long MaterialId;
            public string MaterialCaption;
            public string MaterialCode;
            public long Id;
            public long ObjectId;
            public string Caption;
            public string ObjectCaption;
            public Guid ObjectGuid;

            public int ObjectType;


            /// <summary>
            /// Связь с первым вхождением в сборку; количество элемента из ближайшей связи
            /// </summary>
            public Dictionary<Relation, MeasuredValue> KolvoInAsm
            {
                get
                {
                    if (_kolvoInAsm.Keys.Count == 0)
                    {
                        foreach (var rel in ParentItems)
                        {
                            if (rel.Parent.ObjectType == MetaDataHelper.GetObjectType(new Guid(SystemGUIDs.objtypeAssemblyUnit)).ObjectTypeID)
                                _kolvoInAsm.Add(rel, rel.Kolvo);
                            else
                            {
                                var nextAsms = rel.Parent.KolvoInAsm;
                                foreach (var na in nextAsms)
                                {
                                    _kolvoInAsm[na.Key] = rel.Kolvo;
                                }
                                //_firstAsmsEntersIn.AddRange(rel.Parent.FirstAsmsEntersIn);
                            }
                        }
                    }
                    return _kolvoInAsm;
                }
            }

            //   public MeasuredValue Kolvo;
            /// <summary>
            /// Связи с дочерними объектами
            /// </summary>
            public List<Relation> ChildItems = new List<Relation>();
            /// <summary>
            /// Связи с родительскими объектами
            /// </summary>
            public List<Relation> ParentItems = new List<Relation>();
            public bool HasEmptyKolvo = false;

            /// <summary>
            /// Количество вхождений item в СЕ(в кортеже с количеством)
            /// </summary>
            /// <returns></returns>
            /// 
            private Dictionary<Relation, MeasuredValue> _kolvoInAsm = new Dictionary<Relation, MeasuredValue>();

            public IDictionary<long, MeasuredValue> GetKolvo(bool checkContextObject, ref bool hasContextObject, ref bool hasemptyKolvoRelations)
            {
                MeasuredValue measuredValue = null;
                IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
                if (this.isContextObject) hasContextObject = true;
                foreach (Relation relation in ParentItems)
                {
                    bool hasContextObject1 = true;
                    if (checkContextObject)
                        hasContextObject1 = false;
                    IDictionary<long, MeasuredValue> itemsKolvo = relation.GetKolvo(ref hasContextObject1, ref hasemptyKolvoRelations);

                    hasContextObject = hasContextObject | hasContextObject1;
                    if (checkContextObject && !hasContextObject1 && !this.isContextObject) continue;
                    if (itemsKolvo == null)
                    {
                        continue;
                    }

                    
                    foreach (var itemKolvo in itemsKolvo)
                    {

                        if (!result.TryGetValue(itemKolvo.Key, out measuredValue))
                        {
                            result[itemKolvo.Key] = itemKolvo.Value.Clone() as MeasuredValue;
                            continue;
                        }

                        try
                        {
                            measuredValue.Add(itemKolvo.Value);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(this.ToString() + " " + ex.Message, ex);
                        }
                    }
                }

                return result;
            }

            /// <summary>
            /// Если MaterialId == 0 вернет ObjectId
            /// </summary>
            /// <returns></returns>
            public long GetKey()
            {
                return this.MaterialId != 0
                ? this.MaterialId
                : this.ObjectId;
            }
        }
    }
}