
///Скрипт v.3.1 - Добавлены входимости материалов в ближайшие сборки
//Скрипт v.4.0 - Расчет ведется по технологическому составу
using System;
using System.Windows.Forms;
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
using Intermech.Interfaces.Workflow;

namespace EcoDiffReport
{
    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; private set; }

        public ScriptResult Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            Report sc = new Report();
            sc.Run(session, document, objectIDs);
            return new ScriptResult(true, document);
        }
    }

    public class Report
    {
        //режим расширенного отчета
        private bool compliteReport = true;

        private string _userError = string.Empty;
        private string _adminError = string.Empty;
        private List<long> _messageRecepients = new List<long>() { 3252010 /*Булычева ЕИ*/, 62180376 /*Бородина ЕВ*/ }; //  51448525 /*Протченко ИВ*/
        private long _asm;
        private long _eco;

        private int _partType;
        private int _complectUnitType;
        private int _asmUnitType;
        private int _complectsType;
        private int _zagotType;
        private int _matbaseType;
        private int _sostMaterialType;
        private int _complexPostType;
        private int _MOType;
        private int _CEType;

        //Список идентификаторов версий объектов
        public List<long> contextObjects = new List<long>();

        //Список идентификаторов  объектов
        public List<long> contextObjectsID = new List<long>();

        public List<int> checkKolvoTypes = new List<int>();

        public List<int> enabledTypes = new List<int>();

        /// <summary>
        /// Если не выполняются определенные условия (н-р, не покупное изд), materialid=0, и дальше объект пропускается
        /// </summary>
        /// <param name="row"></param>
        /// <param name="itemsDict">Кортеж словарей в который будут заноситься objectid и Item; linkToObjId и Item. Организует иерархическую связь между Item.</param>
        /// <param name="contextMode"></param>
        /// <returns></returns>
        private Item GetItem(DataRow row, Tuple<Dictionary<long, Item>, Dictionary<long, Item>> itemsDict, bool contextMode)
        {
            Item item = null;

            long linkId = 0;
            object link = row["cad001be-306c-11d8-b4e9-00304f19f545" /*Ссылка на объект*/];
            if (link is long)
                linkId = (long)link;
            long objectId = Convert.ToInt64(row["F_OBJECT_ID"]);

            var itemsDict_byObjId = itemsDict.Item1;
            var itemsDict_byLinkId = itemsDict.Item2;

            if (itemsDict_byObjId.ContainsKey(objectId))
            {
                item = itemsDict_byObjId[objectId];
            }
            else
            {
                item = new Item();
                item.Id = Convert.ToInt64(row["F_PART_ID"]);
                item.ObjectId = Convert.ToInt64(row["F_OBJECT_ID"]);
                item.ObjectType = Convert.ToInt32(row["F_OBJECT_TYPE"]);
                item.LinkToObjId = linkId;
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
            }

            long parentId = Convert.ToInt64(row["F_PROJ_ID"]);
            Relation lnk = new Relation();
            lnk.Child = item;
            lnk.LinkId = Convert.ToInt64(row["F_PRJLINK_ID"]);
            lnk.RelationTypeId = Convert.ToInt32(row["F_RELATION_TYPE"]);

            // если нашли родительский объект в словаре состава, добавляем связи
            if (itemsDict_byObjId.ContainsKey(parentId))
            {
                Item parent = itemsDict_byObjId[parentId];

                if (parent.ObjectType == _sostMaterialType || MetaDataHelper.IsObjectTypeChildOf(parent.ObjectType, _sostMaterialType))
                    return null;

                if (parent.ObjectType == _complexPostType)
                    return null;

                if ((parent.ObjectType == _CEType)
                    && !(new int[] { _MOType, _CEType, _partType }).Contains(item.ObjectType))
                    return null;

                parent.RelationsWithChild.Add(lnk);
                lnk.Parent = parent;
                item.RelationsWithParent.Add(lnk);
            }

            // получаем значение количества и записываем в Link
            object kolvoValue = row["cad00267-306c-11d8-b4e9-00304f19f545"];
            if (kolvoValue is string)
                lnk.Kolvo = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(kolvoValue), false);

            if (lnk.Kolvo == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
            MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/)) != null)
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
            itemsDict_byObjId[objectId] = item;

            if (linkId > 0 && item.ObjectType == _complectUnitType)
            {
                itemsDict_byLinkId[linkId] = item;
            }

            //bool isDopZamen = false;
            //object dopZamenValue = row[Intermech.SystemGUIDs.attributeSubstituteInGroup];
            //if (dopZamenValue != DBNull.Value)
            //{
            //    isDopZamen = Convert.ToInt32(dopZamenValue) != 0;
            //}

            //item.isPocup = isPocup;

            //if (isDopZamen)//Допзамены и их состав не считаем
            //{
            //    AddToLog("isDopZamen " + lnk.ToString());
            //    return null;
            //}

            long materialId = item.Id;

            if (item.ObjectType == _complectUnitType)
            {
                item.MaterialId = materialId;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == _complectsType || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, _complectsType))
            {
                item.MaterialId = materialId;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == _matbaseType || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, _matbaseType))
            {
                item.MaterialId = materialId;
                item.MaterialCaption = item.Caption;
            }

            if (item.ObjectType == _zagotType)
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

                object matId = row["cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/];
                if (matId != DBNull.Value)
                {
                    item.MaterialId = Convert.ToInt64(matId);
                }
            }

            //AddToLog("CreateLink " + lnk.ToString());
            return item;
        }

        private Dictionary<Tuple<long, long>, Item> GetComposition(IDBObject headerObj, DataTable dtCompos, IUserSession session)
        {
            Dictionary<Tuple<long, long>, Item> composition = new Dictionary<Tuple<long, long>, Item>();
            var itemsDict = new Tuple<Dictionary<long, Item>, Dictionary<long, Item>>(new Dictionary<long, Item>(), new Dictionary<long, Item>());

            Item header = new Item();
            header.Id = headerObj.ID;
            header.ObjectId = headerObj.ObjectID;
            header.Caption = headerObj.Caption;
            header.ObjectType = headerObj.ObjectType;
            itemsDict.Item1[header.ObjectId] = header;

            foreach (DataRow row in dtCompos.Rows)
            {
                if (Convert.ToInt32(row["F_OBJECT_TYPE"]) == _zagotType &&
                    row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/].ToString() != "БМЗ")
                    continue;

                var item = GetItem(row, itemsDict, true);
                //if (item != null)
                //    AddToLog("item " + item.ToString());
            }

            itemsDict.Item1.Remove(headerObj.ObjectID);

            foreach (var item in itemsDict.Item1.Values)
            {
                //для собираемой единицы
                if (item.ObjectType == _asmUnitType)
                {
                    // указывает на головную сборку
                    if (item.LinkToObjId == _asm)
                        item.RelationsWithParent = new List<Relation>();

                    //находим комплектующую указывающую на ту же сборку
                    if (itemsDict.Item2.ContainsKey(item.LinkToObjId))
                    {
                        Item complectUnit = itemsDict.Item2[item.LinkToObjId];
                        item.RelationsWithParent = complectUnit.RelationsWithParent;

                        foreach (var relParent in complectUnit.RelationsWithParent)
                        {
                            relParent.Child = item;
                        }
                    }
                }
            }

            foreach (var item in itemsDict.Item1.Values)
            {
                Tuple<string, string> exceptionInfo = null;

                if (item.MaterialId == 0)
                    continue;

                bool hasContextObjects = false;
                var itemsKolvo = item.GetKolvo(true, ref hasContextObjects, ref item.HasEmptyKolvo, ref exceptionInfo);

                if (exceptionInfo.Item2.Length > 0)
                {
                    _userError += string.Format("\nДанные объекта {0} (при формировании отчета для сборки {2} по извещению {3}) введены в систему IPS некорректно и были исключены из отчета. Требуется кооректировка данных. За подробностями обращайтесь к администраторам САПР.\n Сообщение:\n{1}\n", exceptionInfo.Item1, exceptionInfo.Item2, session.GetObject(_asm).Caption, session.GetObject(_eco).Caption);

                    _adminError += string.Format("\nУ пользователя {2} при формировании отчета для сборки {3} по извещению {4} возникла ошибка. Данные объекта {0} введены в систему IPS некорректно и были исключены из отчета. Требуется кооректировка данных.\n Сообщение:\n{1}\n", exceptionInfo.Item1, exceptionInfo.Item2, session.UserName, session.GetObject(_asm).Caption, session.GetObject(_eco).Caption);
                }

                Item mainClone = item.Clone();

                if (!hasContextObjects)
                    continue;

                //if (itemsKolvo == null || itemsKolvo.Count == 0)
                //{
                //    var kolvoItemClone = mainClone.Clone();

                //    Item cachedItem;
                //    var itemKey = new Tuple<long, long>(0, item.GetKey());
                //    if (composition.TryGetValue(itemKey, out cachedItem))
                //    {
                //        if (cachedItem.KolvoSum != null)
                //            cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                //        else
                //            cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                //    }
                //    else
                //    {
                //        composition[itemKey] = kolvoItemClone;
                //    }

                //    AddToLog("res2 " + mainClone);
                //    continue;
                //}

                foreach (var itemKolvo in itemsKolvo)
                {
                    var kolvoItemClone = item.Clone();
                    kolvoItemClone.HasEmptyKolvo = mainClone.HasEmptyKolvo;
                    kolvoItemClone.KolvoSum = itemKolvo.Value;

                    //если несколько заготовок используют одинаковый сортамент, то materialId = id_сортамента. cachedItem в этом случае будет сортамент. К сортаменту добавляется количество материала в заготовке использующей этот сортамент.
                    //ключ состоит из ид.ед.изм и materialid, потому что itemsKolvo для одного item может быть несколько с разными единицами измерения
                    Item cachedItem;
                    var itemKey = new Tuple<long, long>(itemKolvo.Key, item.GetKey());
                    if (composition.TryGetValue(itemKey, out cachedItem))
                    {
                        if (cachedItem.KolvoSum != null)
                            cachedItem.KolvoSum.Add(kolvoItemClone.KolvoSum);
                        else
                            cachedItem.KolvoSum = kolvoItemClone.KolvoSum;
                    }
                    else
                    {
                        composition[itemKey] = kolvoItemClone;
                    }
                }
            }
            return composition;
        }

        public bool Run(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            _asm = objectIDs[0];
            _eco = session.EditingContextID;
            document.Designation = "Отчет сравнения составов";
            if (System.Diagnostics.Debugger.IsAttached)
                System.Diagnostics.Debugger.Break();
            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
            if (System.IO.File.Exists(file))
                System.IO.File.Delete(file);
            //AddToLog("Запускаем скрипт v 4.0");
            _zagotType = MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545" /*Заготовка*/);
            _asmUnitType = MetaDataHelper.GetObjectTypeID("cad00167-306c-11d8-b4e9-00304f19f545" /*Собираемая единица*/);
            _complectUnitType = MetaDataHelper.GetObjectTypeID("cad00166-306c-11d8-b4e9-00304f19f545" /*Комплектующая единица*/);
            _matbaseType = MetaDataHelper.GetObjectTypeID("cad00170-306c-11d8-b4e9-00304f19f545"/*Материал базовый*/);
            _sostMaterialType = MetaDataHelper.GetObjectTypeID("cad00173-306c-11d8-b4e9-00304f19f545" /*Составной материал*/);
            _complectsType = MetaDataHelper.GetObjectTypeID("cad0025f-306c-11d8-b4e9-00304f19f545"/*Комплекты*/);
            _complexPostType = MetaDataHelper.GetObjectTypeID("c248b4f0-69c6-4521-866a-f9837afcb0b2" /*Комплекты поставки*/);
            _CEType = MetaDataHelper.GetObjectTypeID("cad00132-306c-11d8-b4e9-00304f19f545" /*Сборочные единицы*/);
            _MOType = MetaDataHelper.GetObjectTypeID("cad0016f-306c-11d8-b4e9-00304f19f545" /*Маршрут обработки*/);
            _partType = MetaDataHelper.GetObjectTypeID("cad00250-306c-11d8-b4e9-00304f19f545" /*Деталь*/);
            //Типы объектов на которые необходимо проверять заполнено ли количество
            checkKolvoTypes.Add(_complectUnitType);
            checkKolvoTypes.Add(_matbaseType);
            checkKolvoTypes.Add(_sostMaterialType);

            enabledTypes.Add(_CEType);
            enabledTypes.Add(_partType);
            enabledTypes.Add(_MOType);
            //enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)));
            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00185-306c-11d8-b4e9-00304f19f545" /*Техпроцесс базовый*/)));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad001ff-306c-11d8-b4e9-00304f19f545" /*Цехозаход*/));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad0017d-306c-11d8-b4e9-00304f19f545" /*Переход*/));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("b3ec04e4-9d56-4494-a57b-766d10cdfe27" /*Группа материалов*/));
            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00170-306c-11d8-b4e9-00304f19f545" /*Материал базовый*/)));
            enabledTypes.Add(_asmUnitType);
            enabledTypes.Add(_complectUnitType);
            enabledTypes.Add(_zagotType);

            string ecoObjCaption = String.Empty;
            IDBObject ecoObj = session.GetObject(session.EditingContextID, false);
            IDBObject asmObj = session.GetObject(_asm, false);
            if (ecoObj != null)
            {
                var attr = ecoObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
                if (attr != null)
                    document.DocumentName = attr.AsString;

                attr = asmObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
                if (attr != null)
                    document.DocumentName += "__" + attr.AsString;

                ecoObjCaption = ecoObj.Caption;
            }

            document.Designation += " " + DateTime.Now.ToString("dd.MM.yyyy, HH:mm:ss");

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

            rels.Add(MetaDataHelper.GetRelationTypeID("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/));
            rels.Add(MetaDataHelper.GetRelationTypeID("cad00023-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));

            List<ColumnDescriptor> columns = new List<ColumnDescriptor>();

            int attrId = (Int32)ObligatoryObjectAttributes.CAPTION;
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.FieldName, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad00020-306c-11d8-b4e9-00304f19f545" /*Наименование*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad001be-306c-11d8-b4e9-00304f19f545" /*Ссылка на объект*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad0038c-306c-11d8-b4e9-00304f19f545" /*Материал*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            if (addBmzFields)
            {
                attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
            }

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

            #endregion Инициализация запроса к БД

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
            Intermech.SystemGUIDs.filtrationBaseVersions, null, enabledTypes);

            // Храним пару ид версии объекта + ид. физической величины
            // те объекты у которых посчитали количество
            // состав по извещен
            Dictionary<Tuple<long, long>, Item> ecoComposition = GetComposition(headerObj, dt, session);

            //AddToLog("Первый состав с извещением " + headerObj.ObjectID.ToString());

            #endregion Первый состав по извещению

            //сохраняем контекст редактирования
            long sessionid = session.EditingContextID;
            session.EditingContextID = 0;

            #region Второй состав без извещения

            dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
            Intermech.SystemGUIDs.filtrationBaseVersions, null, enabledTypes);

            //AddToLog("Второй состав без извещения " + headerObjBase.ObjectID.ToString());
            session.EditingContextID = 0;

            Dictionary<Tuple<long, long>, Item> baseComposition = GetComposition(headerObjBase, dt, session);

            #endregion Второй состав без извещения

            //возвращаем контекст
            session.EditingContextID = sessionid;

            #region Итоговый состав материалов

            List<Material> resultComposition = new List<Material>();

            foreach (var resItem in baseComposition)
            {
                var baseItem = resItem.Value;
                Material mat = new Material();
                mat.MaterialId = baseItem.Id;
                mat.MaterialCaption = baseItem.Caption;
                mat.type = baseItem.ObjectType;
                mat.linkToObj = baseItem.LinkToObjId;
                //AddToLog("item0 " + baseItem.ToString());
                //добавляем базовый состав в результат
                resultComposition.Add(mat);

                mat.HasEmptyKolvo1 = baseItem.HasEmptyKolvo;
                //AddToLog("mat01 " + mat.ToString());
                if (baseItem.KolvoSum != null)
                {
                    //записываем количество из базовой версии
                    mat.Kolvo1 = baseItem.KolvoSum.Value;
                    mat.MeasureId = baseItem.KolvoSum.MeasureID;
                    var descr = MeasureHelper.FindDescriptor(mat.MeasureId);

                    foreach (var kolvoinasm1 in baseItem.KolvoInAsm)
                    {
                        var asmKolvo = kolvoinasm1.Key.Parent.RelationsWithParent.FirstOrDefault() != null ? kolvoinasm1.Key.Parent.RelationsWithParent.First().Kolvo : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                        var _null = MeasureHelper.ConvertToMeasuredValue(Convert.ToString("0 шт"), false);

                        mat.EntersInAsm1[kolvoinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(kolvoinasm1.Value != null ? kolvoinasm1.Value : _null, asmKolvo != null ? asmKolvo : _null);
                    }

                    if (descr != null)
                        mat.EdIzm = descr.ShortName;
                    else
                    {
                        //AddToLog("descr = null  item.KolvoSum = " + baseItem.KolvoSum.Caption + "  MeasureId = " +
                        //mat.MeasureId);
                    }
                }
                else
                {
                    //AddToLog("item.KolvoSum == null");
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

                //// Отдельно обработаем случай с заготовками
                //if (ecoItem == null
                //&& MetaDataHelper.IsObjectTypeChildOf(resItem.Value.ObjectType, _zagotType))
                //{
                //    Tuple<long, long> zagotItem = null;

                //    foreach (var res2Item in ecoComposition)
                //    {
                //        if (res2Item.Key != resItem.Key ||
                //        res2Item.Value.MaterialId != resItem.Value.MaterialId ||
                //        !MetaDataHelper.IsObjectTypeChildOf(resItem.Value.ObjectType, _zagotType))
                //        {
                //            continue;
                //        }

                //        ecoItem = res2Item.Value;
                //        zagotItem = res2Item.Key;
                //    }

                //    if (zagotItem != null)
                //    {
                //        ecoComposition.Remove(zagotItem);
                //    }
                //}

                if (ecoItem != null)
                {
                    //mat.Kolvo2 = MeasureHelper.ConvertToMeasuredValue(item2.KolvoSum, mat.MeasureId).Value;

                    mat.HasEmptyKolvo2 = ecoItem.HasEmptyKolvo;
                    //AddToLog("item02 " + ecoItem.ToString());
                    if (ecoItem.KolvoSum != null)
                    {
                        //AddToLog("item2.KolvoSum != null");
                        mat.Kolvo2 = MeasureHelper.ConvertToMeasuredValue(ecoItem.KolvoSum, mat.MeasureId).Value;

                        foreach (var kolvoinasm1 in ecoItem.KolvoInAsm)
                        {
                            var asmKolvo = kolvoinasm1.Key.Parent.RelationsWithParent.FirstOrDefault() != null ? kolvoinasm1.Key.Parent.RelationsWithParent.First().Kolvo : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                            var _null = MeasureHelper.ConvertToMeasuredValue(Convert.ToString("0 шт"), false);
                            mat.EntersInAsm2[kolvoinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(kolvoinasm1.Value != null ? kolvoinasm1.Value : _null, asmKolvo != null ? asmKolvo : _null);
                        }
                    }
                    else
                    {
                        //AddToLog("item2.KolvoSum == null");
                        mat.HasEmptyKolvo2 = true;
                    }
                }

                //AddToLog("mat1 " + mat.ToString());
            }

            //Добавим те которых не было в первом наборе
            foreach (var item in ecoComposition.Values)
            {
                Material mat = new Material();
                mat.MaterialId = item.Id;
                mat.MaterialCaption = item.Caption;
                mat.type = item.ObjectType;
                mat.linkToObj = item.LinkToObjId;
                resultComposition.Add(mat);

                mat.HasEmptyKolvo2 = item.HasEmptyKolvo;
                if (item.KolvoSum != null)
                {
                    mat.Kolvo2 = item.KolvoSum.Value;
                    mat.MeasureId = item.KolvoSum.MeasureID;
                    var descr = MeasureHelper.FindDescriptor(mat.MeasureId);

                    foreach (var kolvoinasm1 in item.KolvoInAsm)
                    {
                        var asmKolvo = kolvoinasm1.Key.Parent.RelationsWithParent.FirstOrDefault() != null ? kolvoinasm1.Key.Parent.RelationsWithParent.First().Kolvo : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                        var _null = MeasureHelper.ConvertToMeasuredValue(Convert.ToString("0 шт"), false);
                        mat.EntersInAsm2[kolvoinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(kolvoinasm1.Value != null ? kolvoinasm1.Value : _null, asmKolvo != null ? asmKolvo : _null);
                    }

                    if (descr != null)
                        mat.EdIzm = descr.ShortName;
                    else
                    {
                        //AddToLog("descr = null  item.KolvoSum = " + item.KolvoSum.Caption + "  MeasureId = " +
                        //mat.MeasureId);
                    }
                }
                else
                {
                    mat.HasEmptyKolvo2 = true;
                }

                //AddToLog("mat2 " + mat.ToString());
            }

            List<long> resultCompIds = resultComposition
                .Where(e => e.Kolvo1 != e.Kolvo2)
                .Select(e => e.linkToObj)
                .Where(e => e > 0)
                .Distinct()
                .ToList();

            #region Запись значений атрибутов материалов из соответствующих объектов конструкторского состава
            List<Material> resultComposition_tech = new List<Material>();
            if (resultCompIds.Count > 0)
            {
                IDBObjectCollection col = session.GetObjectCollection(-1);

                List<ConditionStructure> conds = new List<ConditionStructure>();
                conds.Add(new ConditionStructure((Int32)ObligatoryObjectAttributes.F_OBJECT_ID, RelationalOperators.In,
                resultCompIds.ToArray(), LogicalOperators.NONE, 0, false));

                columns = new List<ColumnDescriptor>();

                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_ID, AttributeSourceTypes.Auto,
                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_OBJECT_ID,
                AttributeSourceTypes.Object,
                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.CAPTION,
                AttributeSourceTypes.Object,
                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

                columns.Add(new ColumnDescriptor(
                MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/),
                AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));

                ////attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstituteInGroup) /*Номер заменителя в группе*/);
                ////columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
                ////ColumnNameMapping.Guid, SortOrders.NONE, 0));

                if (addBmzFields)
                {
                    attrId = MetaDataHelper.GetAttributeTypeID(new Guid("8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/) /*Признак изготовления*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
                }

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
                    long id = Convert.ToInt64(row["F_OBJECT_ID"]);
                    string name = Convert.ToString(row["CAPTION"]);
                    string code = Convert.ToString(row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/]);

                    bool isPurchased = false;
                    if (row["8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/] is long)
                    {
                        isPurchased = Convert.ToInt32(row["8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/]) != 1 ? true : false;
                    }

                    var materials = resultComposition.Where(x => x.linkToObj == id);
                    foreach (var mat in materials)
                    {
                        if (mat != null && isPurchased)
                        {
                            mat.MaterialCode = code;
                            mat.MaterialCaption = name;

                            resultComposition_tech.Add(mat);
                            //AddToLog("mat3 " + mat.ToString());
                        }
                    }
                }
            }  
            #endregion

            #endregion

            #region Формирование отчета

            // Заполнение шапки
            //AddToLog("fill header");
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

            int N = 0;
            int totalIndex = 0;
            int rowIndex = 0;
            List<Material> reportComp = new List<Material>();

            foreach (var item in resultComposition_tech)
            {
                if (item.Kolvo1 != item.Kolvo2 || item.HasEmptyKolvo1 != item.HasEmptyKolvo2)
                {
                    Material repeatItem = null;
                    //когда есть повторение позиции, объединяем с уже записанной
                    foreach (var repItem in reportComp)
                    {
                        if (repItem.MaterialCaption == item.MaterialCaption && repItem.type == item.type)
                            repeatItem = repItem;
                    }

                    if (repeatItem != null)
                        repeatItem = repeatItem.Combine(item);
                    else
                        reportComp.Add(item);
                }
            }

            reportComp.RemoveAll(e => reportComp.Count(i => e.MaterialCaption == i.MaterialCaption) > 1 && e.type != _complectUnitType);

            foreach (var item in reportComp.OrderBy(e => e.MaterialCaption))
            {
                //AddToLog("createnode " + item.ToString());
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
                            if (value.Item1 == null || value.Item2 == null || item.EntersInAsm1[key].Item1 == null || item.EntersInAsm1[key].Item2 == null)
                            {
                                item.EntersInAsm1.Remove(key);
                                item.EntersInAsm2.Remove(key);
                                continue;
                            }
                            if (value.Item1.Value == item.EntersInAsm1[key].Item1.Value && value.Item2.Value == item.EntersInAsm1[key].Item2.Value)
                            {
                                item.EntersInAsm1.Remove(key);
                                item.EntersInAsm2.Remove(key);
                            }
                        }
                    }
                    item.EntersInAsm1 = item.EntersInAsm1.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
                    item.EntersInAsm2 = item.EntersInAsm2.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);

                    #region Запись "было"

                    if (item.EntersInAsm1.Count > 0 || item.EntersInAsm2.Count > 0)
                    {
                        N++;
                        totalIndex++;

                        rowIndex++;
                        table.AddChildNode(node, false, false);

                        Write(node, "Индекс", N.ToString());
                        Write(node, "Код", item.MaterialCode);
                        Write(node, "Материал", item.MaterialCaption);

                        if (item.Kolvo1 != 0 || !item.HasEmptyKolvo1)
                        {
                            Write(node, "Всего", Math.Round(item.Kolvo1, 3).ToString() + " " + item.EdIzm);

                            if (N == 1)
                            {
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
                                    DocumentTreeNode col2 = node.FindNode("Столбец2");

                                    totalIndex++;
                                    col2.AddChildNode(node2, false, false);
                                    WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm1.Keys.ToList()[j]);
                                    WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item1.Caption);
                                    WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item2.Caption);
                                }
                            }

                            if (N > 1)
                            {
                                DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", totalIndex));
                                if (item.EntersInAsm1.Count != 0)
                                {
                                    WriteFirstAfterParent(row2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm1.First().Key);
                                    WriteFirstAfterParent(row2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm1.First().Value.Item1.Caption);
                                    WriteFirstAfterParent(row2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm1.First().Value.Item2.Caption);
                                }

                                for (int j = 1; j < item.EntersInAsm1.Count; j++)
                                {
                                    DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                    DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", rowIndex));

                                    totalIndex++;
                                    col2.AddChildNode(node2, false, false);
                                    WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm1.Keys.ToList()[j]);
                                    WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item1.Caption);
                                    WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item2.Caption);
                                }
                            }
                        }
                    }

                    #endregion Запись "было"

                    #region Запись "стало"

                    if (item.EntersInAsm1.Count > 0 || item.EntersInAsm2.Count > 0)
                    {
                        totalIndex++;
                        rowIndex++;
                        node = docrow.CloneFromTemplate(true, true);

                        table.AddChildNode(node, false, false);

                        if (item.Kolvo2 != 0 || !item.HasEmptyKolvo2)
                        {
                            Write(node, "Всего", Math.Round(item.Kolvo2, 3).ToString() + " " + item.EdIzm);

                            DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", totalIndex));
                            DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", rowIndex));
                            if (item.EntersInAsm2.Count != 0)
                            {
                                WriteFirstAfterParent(row2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm2.First().Key);
                                WriteFirstAfterParent(row2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm2.First().Value.Item1.Caption);
                                WriteFirstAfterParent(row2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm2.First().Value.Item2.Caption);
                            }

                            for (int j = 1; j < item.EntersInAsm2.Count; j++)
                            {
                                DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                col2.AddChildNode(node2, false, false);
                                totalIndex++;
                                WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm2.Keys.ToList()[j]);
                                WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm2.Values.ToList()[j].Item1.Caption);
                                WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm2.Values.ToList()[j].Item2.Caption);
                            }
                        }
                    }

                    #endregion Запись "стало"
                }
                else
                {
                    N++;
                    table.AddChildNode(node, false, false);

                    Write(node, "Индекс", N.ToString());
                    Write(node, "ДМТО", item.MaterialDMTO);

                    Write(node, "Код", item.MaterialCode);
                    Write(node, "Материал", item.MaterialCaption);
                    if (item.Kolvo1 != 0 || !item.HasEmptyKolvo1)
                        Write(node, "Было", Convert.ToString(Math.Round(item.Kolvo1, 3)));
                    else
                        Write(node, "Было", "-");
                    if (item.Kolvo2 != 0 || !item.HasEmptyKolvo2)
                        Write(node, "Будет", Convert.ToString(Math.Round(item.Kolvo2, 3)));
                    else
                        Write(node, "Будет", "-");
                    var diff = Math.Round(item.Kolvo2 - item.Kolvo1, 3);
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
                    //AddToLog("item.HasEmptyKolvo  " + item.ToString());
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

            //AddToLog("Завершили создание отчета ");

            #endregion Формирование отчета

            if (_adminError.Length > 0)
            {
                IRouterService router = session.GetCustomService(typeof(IRouterService)) as IRouterService;
                router.CreateMessage(session.SessionGUID, _messageRecepients.ToArray(), "Ошибка формирования отчета сравнения составов (было-стало)", _adminError, session.UserID);
            }

            if (_userError.Length > 0)
                MessageBox.Show(_userError, "Ошибка формирования отчета сравнения составов (было-стало)", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            return true;
        }

        //public void AddToLog(string text)
        //{
        //    string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
        //    text = text + Environment.NewLine;
        //    System.IO.File.AppendAllText(file, text);
        //    //AddToOutputView(text);
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

        private class Material
        {
            private string materialCode;
            private string edIzm = "";

            public bool isPurchased = false;
            public long MaterialId;
            public string MaterialCaption;
            public string MaterialDMTO;
            public int type;
            public long linkToObj;

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

            public Material Combine(Material material)
            {
                if (this.MaterialCode.Length > 0)
                    this.MaterialCode = material.MaterialCode;

                this.Kolvo1 += material.Kolvo1;

                this.Kolvo2 += material.Kolvo2;

                foreach (var eia1 in material.EntersInAsm1)
                {
                    if (!this.EntersInAsm1.ContainsKey(eia1.Key))
                        this.EntersInAsm1.Add(eia1.Key, eia1.Value);
                }

                foreach (var eia2 in material.EntersInAsm2)
                {
                    if (!this.EntersInAsm2.ContainsKey(eia2.Key))
                        this.EntersInAsm2.Add(eia2.Key, eia2.Value);
                }
                return this;
            }

            public override string ToString()
            {
                return string.Format("MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; Kolvo1={3}; Kolvo2={4};",
                MaterialId, MaterialCode, MaterialCaption, Kolvo1, Kolvo2);
            }
        }

        /// <summary>
        /// Связь между объектами Item. Через связь записывается исходное значение количества материала
        /// </summary>
        private class Relation
        {
            public long LinkId;
            public int RelationTypeId;
            public Item Child;
            public Item Parent;
            public MeasuredValue Kolvo;
            public bool HasEmptyKolvo = false;

            public IDictionary<long, MeasuredValue> GetKolvo(ref bool hasContextObject, ref bool hasemptyKolvoRelations, ref Tuple<string, string> exceptionInfo)
            {
                IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
                IDictionary<long, MeasuredValue> itemsKolvo = Parent.GetKolvo(false, ref hasContextObject, ref hasemptyKolvoRelations, ref exceptionInfo);
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
                            exceptionInfo = new Tuple<string, string>(this.Child.Caption, ex.Message);
                            return null;
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

                            #endregion Перенесем выполнение кода статического метода в текущее место

                            AddToLogForLink(text);
                        }
                    }

                    if (Child.RelationsWithChild == null || Child.RelationsWithChild.Count == 0)
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

            public override string ToString()
            {
                string res = "Link=" + LinkId.ToString();
                if (Parent != null)
                    res = res + " parent= " + Parent.Caption.ToString();
                if (Child != null)
                    res = res + " child= " + Child.Caption.ToString();
                res = res + " Количество = " + Convert.ToString(Kolvo);
                return res;
            }

            public void AddToLogForLink(string text)
            {
                string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
                text = text + Environment.NewLine;
                System.IO.File.AppendAllText(file, text);
                //AddToOutputView(text);
            }
        }

        /// <summary>
        /// Объект. Хранит ссылки на связи Link; сумму количества материала (получено из связей)
        /// </summary>
        private class Item
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
                clone.RelationsWithParent = RelationsWithParent;
                clone.RelationsWithChild = RelationsWithChild;
                clone._kolvoInAsm = _kolvoInAsm;
                clone.LinkToObjId = LinkToObjId;
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
            public long LinkToObjId;
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
                        foreach (var rel in RelationsWithParent)
                        {
                            if (rel.Parent.ObjectType == MetaDataHelper.GetObjectType(new Guid("cad00167-306c-11d8-b4e9-00304f19f545" /*Собираемая единица*/)).ObjectTypeID)
                            {
                                _kolvoInAsm.Add(rel, rel.Kolvo);
                                return _kolvoInAsm;
                            }

                            if (rel.Parent.ObjectType == MetaDataHelper.GetObjectType(new Guid(SystemGUIDs.objtypeAssemblyUnit)).ObjectTypeID)
                                _kolvoInAsm.Add(rel, rel.Kolvo);
                            else
                            {
                                var nextAsms = rel.Parent.KolvoInAsm;
                                foreach (var na in nextAsms)
                                {
                                    _kolvoInAsm[na.Key] = rel.Kolvo;
                                }
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
            public List<Relation> RelationsWithChild = new List<Relation>();

            /// <summary>
            /// Связи с родительскими объектами
            /// </summary>
            public List<Relation> RelationsWithParent = new List<Relation>();

            public bool HasEmptyKolvo = false;

            private Dictionary<Relation, MeasuredValue> _kolvoInAsm = new Dictionary<Relation, MeasuredValue>();

            public IDictionary<long, MeasuredValue> GetKolvo(bool checkContextObject, ref bool hasContextObject, ref bool hasemptyKolvoRelations, ref Tuple<string, string> exceptionInfo)
            {
                MeasuredValue measuredValue = null;
                IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
                if (this.isContextObject) hasContextObject = true;

                if (exceptionInfo == null)
                    exceptionInfo = new Tuple<string, string>(this.Caption, string.Empty);

                foreach (Relation relation in RelationsWithParent)
                {
                    bool hasContextObject1 = true;
                    if (checkContextObject)
                        hasContextObject1 = false;
                    IDictionary<long, MeasuredValue> itemsKolvo = relation.GetKolvo(ref hasContextObject1, ref hasemptyKolvoRelations, ref exceptionInfo);

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
                            exceptionInfo = new Tuple<string, string>(this.Caption, ex.Message);
                            return null;
                        }
                    }
                }

                return result;
            }

            /// <summary>
            /// Если MaterialId == 0 вернет ObjectId, в противном случае -- MaterialId.
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