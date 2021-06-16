﻿///Скрипт v.2.1 от 28.09.2020 Убраны все static для перехода на IPS 6

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

    public class Script10
    {
        public bool Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            //try
            {
                document.Designation = "Сравнение составов";
                if (System.Diagnostics.Debugger.IsAttached)
                    System.Diagnostics.Debugger.Break();
                string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
                if (System.IO.File.Exists(file))
                    System.IO.File.Delete(file);
                AddToLog("Запускаем скрипт v 2.2");
                //throw new Exception("jhghjhv");
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
                    attrId = MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));

                    //attrId = MetaDataHelper.GetAttributeTypeID( "b1e25726-587e-4bb9-8934-75b85b57d5eb" /*Ответственный за номенклатуру ДМТО (БМЗ)*/);
                    //columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    //ColumnNameMapping.Guid, SortOrders.NONE, 0));

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

                //Первый состав
                DataTable dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
                Intermech.SystemGUIDs.filtrationBaseVersions, null);




                // Храним пару ид. версии объекта + ид. физической величины
                // те объекты у которых посчитали количество

                Dictionary<Tuple<long, long>, Item> res2 = new Dictionary<Tuple<long, long>, Item>();
                AddToLog("Первый состав с извещением " + headerObj.ObjectID.ToString());
                if (dt != null)
                {
                    //dt.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script1.log");
                    Dictionary<long, Item> itemsDict = new Dictionary<long, Item>();
                    Item header = new Item();
                    header.Id = headerObj.ID;
                    header.ObjectId = headerObj.ObjectID;
                    header.Caption = headerObj.Caption;
                    header.ObjectType = headerObj.ObjectType;
                    itemsDict[header.ObjectId] = header;
                    foreach (DataRow row in dt.Rows)
                    {
                        var item = GetItem(row, itemsDict, true);
                        if (item != null)
                            AddToLog("item " + item.ToString());
                    }

                    itemsDict.Remove(headerObj.ObjectID);
                    foreach (var item in itemsDict.Values)
                    {
                        if (item.MaterialId != 0)
                        {
                            Item mainClone = item.Clone();
                            bool hasContextObjects = false;
                            var itemsKolvo = item.GetKolvo(true, ref hasContextObjects, ref mainClone.HasEmptyKolvo);
                            if (!hasContextObjects)
                                continue;
                            if (itemsKolvo == null || itemsKolvo.Count == 0)
                            {
                                var kolvoItemClone = mainClone.Clone();

                                Item cachedItem;
                                var itemKey = new Tuple<long, long>(0, item.GetKey());
                                if (res2.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    res2[itemKey] = kolvoItemClone;
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
                                if (res2.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    res2[itemKey] = kolvoItemClone;
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

                long sessionid = session.EditingContextID;
                //Второй состав без извещения
                AddToLog("Второй состав без извещения" + headerObjBase.ObjectID.ToString());
                session.EditingContextID = 0;
                items = new List<ObjInfoItem>();
                items.Add(new ObjInfoItem(headerObjBase.ObjectID));
                dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
                Intermech.SystemGUIDs.filtrationBaseVersions, null);

                Dictionary<Tuple<long, long>, Item> res1 = new Dictionary<Tuple<long, long>, Item>();
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
                        var item = GetItem(row, itemsDict, false);
                        if (item != null)
                            AddToLog("item " + item.ToString());
                    }

                    itemsDict.Remove(headerObjBase.ObjectID);
                    foreach (var item in itemsDict.Values)
                    {
                        if (item.MaterialId != 0)
                        {
                            Item mainClone = item.Clone();
                            bool hasContextObjects = false;
                            var itemsKolvo = item.GetKolvo(true, ref hasContextObjects, ref mainClone.HasEmptyKolvo);
                            if (!hasContextObjects)
                                continue;


                            if (itemsKolvo == null || itemsKolvo.Count == 0)
                            {
                                var kolvoItemClone = mainClone.Clone();

                                Item cachedItem;
                                var itemKey = new Tuple<long, long>(0, item.GetKey());
                                if (res1.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    res1[itemKey] = kolvoItemClone;
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
                                if (res1.TryGetValue(itemKey, out cachedItem))
                                {
                                    if (cachedItem.KolvoSum != null)
                                        cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                                    else
                                        cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                                }
                                else
                                {
                                    res1[itemKey] = kolvoItemClone;
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

                session.EditingContextID = sessionid;
                AddToLog("q ");
                List<Material> res = new List<Material>();
                foreach (var resItem in res1)
                {
                    var item = resItem.Value;
                    Material mat = new Material();
                    if (item.MaterialDMTO != null)
                        mat.MaterialDMTO = item.MaterialDMTO;
                    mat.MaterialId = item.MaterialId;
                    mat.MaterialCode = item.MaterialCode;
                    mat.MaterialCaption = item.MaterialCaption;
                    AddToLog("item0 " + item.ToString());
                    res.Add(mat);
                    mat.HasEmptyKolvo1 = item.HasEmptyKolvo;
                    AddToLog("mat01 " + mat.ToString());
                    if (item.KolvoSum != null)
                    {
                        AddToLog("item.KolvoSum != null");
                        mat.Kolvo1 = item.KolvoSum.Value;
                        mat.MeasureId = item.KolvoSum.MeasureID;
                        var descr = MeasureHelper.FindDescriptor(mat.MeasureId);
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
                        AddToLog("item.KolvoSum == null");
                        mat.HasEmptyKolvo1 = true;
                    }

                    Item item2;
                    if (res2.TryGetValue(resItem.Key, out item2))
                    {
                        res2.Remove(resItem.Key);
                    }
                    else
                    {
                        var emptyItem = new Tuple<long, long>(0, resItem.Key.Item2);
                        if (res2.TryGetValue(emptyItem, out item2))
                        {
                            res2.Remove(emptyItem);
                        }
                    }

                    // Отдельно обработаем случай с заготовками 
                    if (item2 == null
                    && MetaDataHelper.IsObjectTypeChildOf(resItem.Value.ObjectType, zagot))
                    {
                        Tuple<long, long> zagotItem = null;

                        foreach (var res2Item in res2)
                        {
                            if (res2Item.Key.Item1 != resItem.Key.Item1 ||
                            res2Item.Value.MaterialId != resItem.Value.MaterialId ||
                            !MetaDataHelper.IsObjectTypeChildOf(resItem.Value.ObjectType, zagot))
                            {
                                continue;
                            }

                            item2 = res2Item.Value;
                            zagotItem = res2Item.Key;
                        }

                        if (zagotItem != null)
                        {
                            res2.Remove(zagotItem);
                        }
                    }


                    if (item2 != null)
                    {
                        //mat.Kolvo2 = MeasureHelper.ConvertToMeasuredValue(item2.KolvoSum, mat.MeasureId).Value;
                        if (item2.MaterialDMTO != null)
                            mat.MaterialDMTO = item2.MaterialDMTO;
                        mat.HasEmptyKolvo2 = item2.HasEmptyKolvo;
                        AddToLog("item02 " + item2.ToString());
                        if (item2.KolvoSum != null)
                        {
                            AddToLog("item2.KolvoSum != null");
                            mat.Kolvo2 = MeasureHelper.ConvertToMeasuredValue(item2.KolvoSum, mat.MeasureId).Value;
                        }
                        else
                        {
                            AddToLog("item2.KolvoSum == null");
                            mat.HasEmptyKolvo2 = true;
                        }

                    }

                    AddToLog("mat1 " + mat.ToString());
                }

                foreach (var item in res2.Values) //Добавим те которых не было в первом наборе
                {
                    Material mat = new Material();
                    if (item.MaterialDMTO != null)
                        mat.MaterialDMTO = item.MaterialDMTO;
                    mat.MaterialDMTO = item.MaterialDMTO;
                    mat.MaterialId = item.MaterialId;
                    mat.MaterialCode = item.MaterialCode;
                    mat.MaterialCaption = item.MaterialCaption;
                    res.Add(mat);
                    mat.HasEmptyKolvo2 = item.HasEmptyKolvo;
                    if (item.KolvoSum != null)
                    {
                        mat.Kolvo2 = item.KolvoSum.Value;
                        mat.MeasureId = item.KolvoSum.MeasureID;
                        var descr = MeasureHelper.FindDescriptor(mat.MeasureId);
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
                foreach (var item in res)
                {
                    ids.Add(item.MaterialId);
                    if (string.IsNullOrEmpty(item.MaterialCode))
                    {

                    }
                }

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
                        Material mat = res.Find(x => x.MaterialId == id);
                        if (mat != null)
                        {
                            mat.MaterialCode = code;
                            mat.MaterialCaption = name;
                            AddToLog("mat3 " + mat.ToString());
                        }
                    }

                    foreach (var item in res)
                    {
                        if (item.MaterialCode == "")
                        {
                            item.MaterialCode = "";
                            IDBObject obj = session.GetObject(item.MaterialId, false);
                            if (obj != null)
                            {
                                IMSObjectType type = MetaDataHelper.GetObjectType(obj.ObjectType);
                                IDBAttribute attr =
                                obj.GetAttributeByGuid(new Guid("cad00020-306c-11d8-b4e9-00304f19f545"), false);
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



                foreach (var item in res)
                {

                    AddToLog("beforesort " + item.ToString());
                }


                AddToLog("res.Sort");

                res.Sort(Compare);

                foreach (var item in res)
                {

                    AddToLog("aftersort " + item.ToString());
                }




                List<Material> restemp = new List<Material>();
                foreach (var item in res)
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

                res = restemp;

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
                foreach (var item in res)
                {

                    AddToLog("beforecreatenode " + item.ToString());
                }

                int index = 0;
                foreach (var item in res)
                {
                    index++;
                    // for (int i = 0; i < 30; i++)
                    {
                        if (item.Kolvo1 != item.Kolvo2 || item.HasEmptyKolvo1 != item.HasEmptyKolvo2)
                        {
                            AddToLog("createnode " + item.ToString());
                            DocumentTreeNode node = docrow.CloneFromTemplate(true, true);
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
            }
            //catch (Exception ex)
            //{
            //AddToLog(ex.Message);
            //AddToLog(ex.StackTrace);
            //throw ex;
            //}

            //Вставьте ваш код сценария здесь
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

        private void Write(DocumentTreeNode parent, string tplid, string text)
        {
            TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
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
            
            //76468752 for "3ТЭ25К2М.001.11.030 (Воздуховод правый)" 
            long parentId = Convert.ToInt64(row["F_PROJ_ID"]);
            Link lnk = new Link();
            lnk.Child = item;
            lnk.LinkId = Convert.ToInt64(row["F_PRJLINK_ID"]); //-76468757
            lnk.RelationTypeId = Convert.ToInt32(row["F_RELATION_TYPE"]);
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

            if (isDopZamen)//Допзамены и их состав не считаем
            {
                AddToLog("isDopZamen " + lnk.ToString());
                return null;
            }

            itemsDict[id] = item;



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
            public double Kolvo1 = 0;
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

            public override string ToString()
            {
                return string.Format("MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; Kolvo1={3}; Kolvo2={4};",
                MaterialId, MaterialCode, MaterialCaption, Kolvo1, Kolvo2);
            }
        }

        class Link
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
                                result[MeasureHelper.FindDescriptor(itemKolvo.Value).PhysicalQuantityID] = itemKolvo.Value;
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
                //     clone.Kolvo = Kolvo;
                return clone;
            }

            public string MaterialDMTO = null;
            public bool isContextObject = false;
            public bool isPocup = false;
            public MeasuredValue KolvoSum;
            public long MaterialId;
            public string MaterialCaption;
            public string MaterialCode;
            public long Id;
            public long ObjectId;
            public string Caption;
            public string ObjectCaption;
            public Guid ObjectGuid;

            public int ObjectType;

            //   public MeasuredValue Kolvo;
            public List<Link> ChildItems = new List<Link>();
            public List<Link> ParentItems = new List<Link>();
            public bool HasEmptyKolvo = false;

            public IDictionary<long, MeasuredValue> GetKolvo(bool checkContextObject, ref bool hasContextObject, ref bool hasemptyKolvoRelations)
            {
                long phId;
                MeasuredValue measuredValue = null;
                IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
                if (this.isContextObject) hasContextObject = true;
                foreach (Link item in ParentItems)
                {
                    bool hasContextObject1 = true;
                    if (checkContextObject)
                        hasContextObject1 = false;
                    IDictionary<long, MeasuredValue> itemsKolvo = item.GetKolvo(ref hasContextObject1, ref hasemptyKolvoRelations);
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

            public long GetKey()
            {
                return this.MaterialId != 0
                ? this.MaterialId
                : this.ObjectId;
            }
        }
    }
}
