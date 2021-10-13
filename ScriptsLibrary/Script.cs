////Скрипт v.3.1 - Добавлены входимости материалов в ближайшие сборки
////Скрипт v.4.0 - Расчет ведется по технологическому составу
//using System;
//using System.Windows.Forms;
//using System.Linq;
//using System.Xml;
//using Intermech.Interfaces;
//using Intermech.Interfaces.Document;
//using Intermech.Expert;
//using Intermech.Kernel.Search;
//using System.Data;
//using Intermech.Document.Model;
//using System.Collections.Generic;
//using Intermech.Interfaces.Compositions;
//using System.Collections.Specialized;
//using System.Collections;
//using System.Drawing;
//using Intermech.Expert.Scenarios;
//using Intermech.Interfaces.Client;
//using Intermech.Navigator.Interfaces.QuickSearch;
//using Intermech.Interfaces.Contexts;
//using Intermech.Collections;
//using Intermech;
//using Intermech.Interfaces.Workflow;
//namespace EcoDiffReport
//{
//    public class Script
//    {
//        public ICSharpScriptContext ScriptContext { get; private set; }

//        public ScriptResult Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
//        {
//            Report report = new Report()
//            {
//                compliteReport = true /*Устанавливаем режим отчета: true - расширенный, false - обычный*/,
//                originOrg = "БМЗ"
//            };
//            report.Run(session, document, objectIDs);

//            if (report.compliteReport)
//                document.UpdateLayout(true);

//            return new ScriptResult(true, document);
//        }
//    }

//    public class Report
//    {
//        public bool compliteReport = false;

//        public bool writeLog = false;
//        public bool writeTestingData = false;

//        private string _userError = string.Empty;
//        private string _adminError = string.Empty;
//        private List<long> _messageRecepients = new List<long>();
//        //private List<long> _messageRecepients = new List<long>() { 3252010 /*Булычева ЕИ*/, 62180376 /*Бородина ЕВ*/ }; //  51448525 /*Протченко ИВ*/
//        private long _asm;
//        private long _eco;

//        private int _partType;
//        private int _complectUnitType;
//        private int _asmUnitType;
//        private int _complectsType;
//        private int _zagotType;
//        private int _matbaseType;
//        private int _sostMaterialType;
//        private int _complexPostType;
//        private int _MOType;
//        private int _CEType;

//        //Список идентификаторов версий объектов
//        public List<long> contextObjects = new List<long>();

//        //Список идентификаторов  объектов
//        public List<long> contextObjectsID = new List<long>();

//        public List<int> checkAmountTypes = new List<int>();

//        public List<int> enabledTypes = new List<int>();

//        public string originOrg;

//        /// <summary>
//        /// Если не выполняются определенные условия (н-р, не покупное изд), materialid=0, и дальше объект пропускается
//        /// </summary>
//        /// <param name="row"></param>
//        /// <param name="itemsDict">Кортеж словарей в который будут заноситься objectid и Item; linkToObjId и Item. Организует иерархическую связь между Item.</param>
//        /// <param name="contextMode"></param>
//        /// <returns></returns>
//        private Item GetItem(DataRow row, Tuple<Dictionary<long, Item>, Dictionary<long, Item>> itemsDict, bool contextMode, CompositionType compositionType)
//        {
//            Item item = null;

//            long linkId = 0;
//            object link = row["cad001be-306c-11d8-b4e9-00304f19f545" /*Ссылка на объект*/];
//            if (link is long)
//                linkId = (long)link;
//            long objectId = Convert.ToInt64(row["F_OBJECT_ID"]);

//            var itemsDict_byObjId = itemsDict.Item1;
//            var itemsDict_byLinkId = itemsDict.Item2;

//            if (itemsDict_byObjId.ContainsKey(objectId))
//            {
//                item = itemsDict_byObjId[objectId];
//            }
//            else
//            {
//                var objType = Convert.ToInt32(row["F_OBJECT_TYPE"]);

//                if (objType == _sostMaterialType)
//                    item = new ComplexMaterialItem();
//                else if (objType == _matbaseType || MetaDataHelper.IsObjectTypeChildOf(objType, _matbaseType))
//                    item = new MaterialItem();
//                else
//                    item = new Item();

//                item.SourseOrg = Convert.ToString(row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/]);
//                item.Id = Convert.ToInt64(row["F_PART_ID"]);
//                item.ObjectId = Convert.ToInt64(row["F_OBJECT_ID"]);
//                item.ObjectType = objType;
//                item.LinkToObjId = linkId;
//                string designation = Convert.ToString(row[Intermech.SystemGUIDs.attributeDesignation]);
//                string name = Convert.ToString(row[Intermech.SystemGUIDs.attributeName]);
//                string caption = name;
//                if (!string.IsNullOrEmpty(designation))
//                    caption = designation + " " + name;
//                item.Caption = caption;
//                if (!contextMode)
//                {
//                    if (contextObjects.Contains(item.ObjectId) || contextObjects.Contains(-item.ObjectId)) return null; //Объект входит в состав извещения значит считаем что в составе без извещения его нет
//                }
//                if (contextObjectsID.Contains(item.Id))
//                    item.isContextObject = true;

//                //if (item is ComplexMaterialItem)
//                //{
//                //    var Amount1 = row["b9ab8a1a-e988-4031-9f30-72a95644830a" /*Количество компонента 1 (БМЗ)*/];
//                //    if (Amount1 is string)
//                //        (item as ComplexMaterialItem).Component1Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(Amount1), false);

//                //    var Amount2 = row["8de3b657-868a-47de-8a62-3bbaf7c4994f" /*Количество компонента 2 (БМЗ)*/];
//                //    if (Amount2 is string)
//                //        (item as ComplexMaterialItem).Component2Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(Amount2), false);

//                //    var AmountMain = row["34b64b00-d430-48e4-8692-f52a7485a10a" /*Количество основы (БМЗ)*/];
//                //    if (AmountMain is string)
//                //        (item as ComplexMaterialItem).MainComponentAmount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(AmountMain), false);
//                //}
//                //if (item is MaterialItem)
//                //{
//                //    var compSostMat = row["54b2e64e-1d4b-4a9d-9633-023f6b6bcd4a" /*Компонент составного материала*/];
//                //    if (compSostMat is string)
//                //    {
//                //        switch (compSostMat.ToString())
//                //        {
//                //            case "Основа":
//                //                (item as MaterialItem).ComplexMaterialComponentValue = MaterialItem.ComplexMaterialComponent.ComponentMain;
//                //                break;

//                //            case "Компонент 1":
//                //                (item as MaterialItem).ComplexMaterialComponentValue = MaterialItem.ComplexMaterialComponent.Component1;
//                //                break;

//                //            case "Компонент 2":
//                //                (item as MaterialItem).ComplexMaterialComponentValue = MaterialItem.ComplexMaterialComponent.Component2;
//                //                break;

//                //            default:
//                //                break;
//                //        }
//                //    }
//                //}
//            }

//            long parentId = Convert.ToInt64(row["F_PROJ_ID"]);
//            Relation lnk = new Relation();
//            lnk.Child = item;
//            lnk.LinkId = Convert.ToInt64(row["F_PRJLINK_ID"]);
//            lnk.RelationTypeId = Convert.ToInt32(row["F_RELATION_TYPE"]);

//            // если нашли родительский объект в словаре состава, добавляем связи
//            if (itemsDict_byObjId.ContainsKey(parentId))
//            {
//                Item parent = itemsDict_byObjId[parentId];

//                if (parent.ObjectType == _complexPostType)
//                    return null;

//                parent.RelationsWithChild.Add(lnk);
//                lnk.Parent = parent;
//                item.RelationsWithParent.Add(lnk);
//            }
//            object AmountValue = null;
//            //для составного не записываем количество, чтобы в расчет составных материалов попадала только се
//            if (item.ObjectType != _sostMaterialType)
//            {
//                // получаем значение количества и записываем в Link
//                AmountValue = row["cad00267-306c-11d8-b4e9-00304f19f545"];
//                if (AmountValue is string)
//                    lnk.Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(AmountValue), false);
//            }

//            if (lnk.Amount == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
//            MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/)) != null)
//            {
//                foreach (var checkAmountType in checkAmountTypes)
//                {
//                    if (lnk.Child != null && (lnk.Child.ObjectType == checkAmountType ||
//                    MetaDataHelper.IsObjectTypeChildOf(lnk.Child.ObjectType, checkAmountType)))
//                    {
//                        lnk.HasEmptyAmount = true;
//                    }
//                }
//            }
//            itemsDict_byObjId[objectId] = item;

//            //для комплектующих единиц записываем группы замен
//            if (item.ObjectType == _complectUnitType)
//            {
//                object substituteInGroup = row[Intermech.SystemGUIDs.attributeSubstituteInGroup];
//                if (substituteInGroup != DBNull.Value)
//                {
//                    int groupNo = Convert.ToInt32(substituteInGroup);
//                    item.isActualReplacement = groupNo == 0;
//                    item.isPossableReplacement = groupNo > 0;
//                }
//                object substitutesGroupNo = row[Intermech.SystemGUIDs.attributeSubstitutesGroupNo];
//                if (substitutesGroupNo != DBNull.Value)
//                {
//                    item.ReplacementGroup = Convert.ToInt32(substitutesGroupNo);
//                }
//            }

//            if (linkId > 0 && item.ObjectType == _complectUnitType)
//            {
//                itemsDict_byLinkId[linkId] = item;
//            }

//            object isPocupValue = row["8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/];
//            if (isPocupValue != DBNull.Value)
//            {
//                item.isPurchased = Convert.ToInt32(isPocupValue) != 1;
//                item.isCoop = Convert.ToInt32(isPocupValue) == 3;
//            }

//            object matCode = row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/];
//            if (matCode != DBNull.Value)
//            {
//                item.MaterialCode = Convert.ToString(matCode);
//            }

//            long materialId = item.Id;

//            if (item.ObjectType == _complectUnitType)
//            {
//                item.MaterialId = materialId;
//            }

//            if (item.ObjectType == _complectsType || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, _complectsType))
//            {
//                item.MaterialId = materialId;
//            }

//            if (item.ObjectType != _sostMaterialType && (item.ObjectType == _matbaseType || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, _matbaseType)))
//            {
//                item.MaterialId = materialId;
//            }

//            if (item.ObjectType == _zagotType)
//            {
//                AmountValue = row["cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/];
//                if (AmountValue is string)
//                {
//                    lnk.Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(AmountValue), false);
//                }

//                if (lnk.Amount == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
//                MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/)) != null)
//                {
//                    lnk.HasEmptyAmount = true;
//                }

//                object matId = row["cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/];
//                if (matId != DBNull.Value)
//                {
//                    item.MaterialId = Convert.ToInt64(matId);
//                }
//            }
//            //if (item.ObjectType == _asmUnitType || item.ObjectType == _CEType)
//            //    item.MaterialId = materialId;

//            //AddToLog("CreateLink " + lnk.ToString());
//            return item;
//        }
//        public enum CompositionType
//        {
//            construction = 1,
//            technology = 2
//        }

//        private Dictionary<Tuple<long, long, string>, Item> GetComposition(IDBObject headerObj, DataTable dtCompos, IUserSession session, CompositionType compositionType = CompositionType.technology)
//        {
//            Dictionary<Tuple<long, long, string>, Item> composition = new Dictionary<Tuple<long, long, string>, Item>();
//            var itemsDict = new Tuple<Dictionary<long, Item>, Dictionary<long, Item>>(new Dictionary<long, Item>(), new Dictionary<long, Item>());

//            Item header = new Item();
//            header.Id = headerObj.ID;
//            header.ObjectId = headerObj.ObjectID;
//            header.Caption = headerObj.Caption;
//            header.ObjectType = headerObj.ObjectType;
//            itemsDict.Item1[header.ObjectId] = header;

//            foreach (DataRow row in dtCompos.Rows)
//            {
//                Item item = GetItem(row, itemsDict, true, compositionType);
//            }

//            itemsDict.Item1.Remove(headerObj.ObjectID);

//            foreach (var item in itemsDict.Item1.Values)
//            {
//                //для собираемой единицы
//                if (item.ObjectType == _asmUnitType)
//                {
//                    // указывает на головную сборку
//                    if (item.LinkToObjId == _asm)
//                        item.RelationsWithParent = new List<Relation>();

//                    Item complectUnit = null;
//                    //находим комплектующую указывающую на ту же сборку
//                    if (itemsDict.Item2.TryGetValue(item.LinkToObjId, out complectUnit))
//                    {
//                        item.RelationsWithParent = complectUnit.RelationsWithParent;

//                        foreach (var relParent in complectUnit.RelationsWithParent)
//                        {
//                            relParent.Child = item;
//                        }
//                    }
//                }

//                //обновляем количество сборок и деталей (констр. сост.) из данных технологического состава
//                if (item.ObjectType == _partType || item.ObjectType == _CEType)
//                {
//                    //(по ссылке на объект не удается связать комплектующую со сборкой или деталью, -- ссылка указывает на ид версии базового объекта, а не текущего)
//                    //если комплектующая associatedComplectUnit входит в собираемую с тем же названием, что и item в вышестоящую сборку
//                    IEnumerable<string> itemEntersInAsms = item.EntersInAsms.Values.Select(e => e.Caption);

//                    Item associatedComplectUnit = itemsDict.Item2.Values
//                        .Where(e => e.ObjectType == _complectUnitType)
//                        .Where(e => e.Caption == item.Caption)
//                        .FirstOrDefault(e => e.EntersInAsms.Values.Select(r => r.Caption).Where(t => itemEntersInAsms.Contains(t)).Count() > 0);

//                    //связываем итоговое количество
//                    if (associatedComplectUnit != null)
//                        item.AmountSum = associatedComplectUnit.AmountSum;

//                    //перезаписываем значения количества для item из данных associatedComplectUnit
//                    foreach (Relation item_rwp in item.RelationsWithParent)
//                    {
//                        if (associatedComplectUnit == null)
//                            break;

//                        Relation associatedItem_rwp = associatedComplectUnit.RelationsWithParent
//                            .FirstOrDefault(e => e.Child.Caption == associatedComplectUnit.Caption);

//                        item_rwp.Amount = associatedItem_rwp != null ? associatedItem_rwp.Amount : item_rwp.Amount;
//                    }
//                }

//                //if (item is ComplexMaterialItem)
//                //{
//                //    ComplexMaterialItem complexMaterial = item as ComplexMaterialItem;

//                //    foreach (var item_rwc in item.RelationsWithChild)
//                //    {
//                //        switch ((item_rwc.Child as MaterialItem).ComplexMaterialComponentValue)
//                //        {
//                //            case MaterialItem.ComplexMaterialComponent.Component1:
//                //                item_rwc.Amount = item_rwc.Amount == null ? complexMaterial.Component1Amount : item_rwc.Amount;
//                //                break;

//                //            case MaterialItem.ComplexMaterialComponent.Component2:
//                //                item_rwc.Amount = item_rwc.Amount == null ? complexMaterial.Component2Amount : item_rwc.Amount;
//                //                break;

//                //            case MaterialItem.ComplexMaterialComponent.ComponentMain:
//                //                item_rwc.Amount = item_rwc.Amount == null ? complexMaterial.MainComponentAmount : item_rwc.Amount;
//                //                break;

//                //            default:
//                //                break;
//                //        }
//                //    }
//                //}
//            }

//            foreach (Item item in itemsDict.Item1.Values)
//            {
//                Tuple<string, string> exceptionInfo = null;

//                if (item.MaterialId == 0)
//                    continue;

//                bool hasContextObjects = false;

//                var itemsAmount = item.GetAmount(true, ref hasContextObjects, ref item.HasEmptyAmount, exceptionInfo, null);

//                if (exceptionInfo != null)
//                {
//                    if (exceptionInfo.Item2.Length > 0)
//                    {
//                        _userError += string.Format("\nДанные объекта {0} (при формировании отчета для сборки {2} по извещению {3}) введены в систему IPS некорректно и были исключены из отчета. Требуется кооректировка данных. За подробностями обращайтесь к администраторам САПР.\n Сообщение:\n{1}\n", exceptionInfo.Item1, exceptionInfo.Item2, session.GetObject(_asm).Caption, session.GetObject(_eco).Caption);

//                        _adminError += string.Format("\nУ пользователя {2} при формировании отчета для сборки {3} по извещению {4} возникла ошибка. Данные объекта {0} введены в систему IPS некорректно и были исключены из отчета. Требуется кооректировка данных.\n Сообщение:\n{1}\n", exceptionInfo.Item1, exceptionInfo.Item2, session.UserName, session.GetObject(_asm).Caption, session.GetObject(_eco).Caption);
//                    }
//                }

//                Item mainClone = item.Clone();

//                if (!hasContextObjects)
//                    continue;

//                foreach (var itemAmount in itemsAmount)
//                {
//                    var AmountItemClone = item.Clone();
//                    AmountItemClone.HasEmptyAmount = mainClone.HasEmptyAmount;
//                    AmountItemClone.AmountSum = itemAmount.Value;

//                    //если несколько заготовок используют одинаковый сортамент, то materialId = id_сортамента. cachedItem в этом случае будет сортамент. К сортаменту добавляется количество материала в заготовке использующей этот сортамент.
//                    //ключ состоит из ид.ед.изм и materialid, потому что itemsAmount для одного item может быть несколько с разными единицами измерения

//                    Item actualItem;
//                    Tuple<long, long, string> itemKey = new Tuple<long, long, string>(itemAmount.Key, item.GetKey(), item.SourseOrg);
//                    if (composition.TryGetValue(itemKey, out actualItem))
//                    {
//                        if (actualItem.AmountSum != null)
//                        {
//                            actualItem.AssociatedItemsAndSelf[item.ObjectId] = (AmountItemClone);
//                            actualItem.AmountSum.Add(AmountItemClone.AmountSum);

//                            //обновляем данные по входямостям в сборки
//                            foreach (var newItem in AmountItemClone.EntersInAsms)
//                            {
//                                if (!actualItem.EntersInAsms.ContainsKey(newItem.Key))
//                                    actualItem.EntersInAsms[newItem.Key] = (newItem.Value);
//                            }
//                        }
//                        else
//                            actualItem.AmountSum = AmountItemClone.AmountSum;
//                    }
//                    else
//                    {
//                        composition[itemKey] = AmountItemClone;
//                    }
//                }
//            }

//            //отдельно считаем сборки и собираемые, чтобы не вносить изменения в расчет материалов:
//            foreach (var asm in itemsDict.Item1.Values.Where(e => e.ObjectType == _asmUnitType || e.ObjectType == _CEType
//            //|| MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)).Contains(e.ObjectType)
//            ))
//            {
//                bool hasContextObjects = false;
//                asm.AmountSum = asm.GetAmount(false, ref hasContextObjects, ref asm.HasEmptyAmount, null, null).Values.FirstOrDefault();
//            }

//            return composition;
//        }

//        public bool Run(IUserSession session, ImDocumentData document, Int64[] objectIDs)
//        {
//            _asm = objectIDs[0];
//            _eco = session.EditingContextID;
//            document.Designation = "Отчет сравнения составов";

//            if (System.Diagnostics.Debugger.IsAttached)
//                System.Diagnostics.Debugger.Break();

//            string logFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
//            if (System.IO.File.Exists(logFile))
//                System.IO.File.Delete(logFile);

//            string testFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\testData.txt";
//            if (System.IO.File.Exists(testFile))
//                System.IO.File.Delete(testFile);

//            //AddToLog("Запускаем скрипт v 4.0");
//            _zagotType = MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545" /*Заготовка*/);
//            _asmUnitType = MetaDataHelper.GetObjectTypeID("cad00167-306c-11d8-b4e9-00304f19f545" /*Собираемая единица*/);
//            _complectUnitType = MetaDataHelper.GetObjectTypeID("cad00166-306c-11d8-b4e9-00304f19f545" /*Комплектующая единица*/);
//            _matbaseType = MetaDataHelper.GetObjectTypeID("cad00170-306c-11d8-b4e9-00304f19f545"/*Материал базовый*/);
//            _sostMaterialType = MetaDataHelper.GetObjectTypeID("cad00173-306c-11d8-b4e9-00304f19f545" /*Составной материал*/);
//            _complectsType = MetaDataHelper.GetObjectTypeID("cad0025f-306c-11d8-b4e9-00304f19f545"/*Комплекты*/);
//            _complexPostType = MetaDataHelper.GetObjectTypeID("c248b4f0-69c6-4521-866a-f9837afcb0b2" /*Комплекты поставки*/);
//            _CEType = MetaDataHelper.GetObjectTypeID("cad00132-306c-11d8-b4e9-00304f19f545" /*Сборочные единицы*/);
//            _MOType = MetaDataHelper.GetObjectTypeID("cad0016f-306c-11d8-b4e9-00304f19f545" /*Маршрут обработки*/);
//            _partType = MetaDataHelper.GetObjectTypeID("cad00250-306c-11d8-b4e9-00304f19f545" /*Деталь*/);
//            //Типы объектов на которые необходимо проверять заполнено ли количество
//            checkAmountTypes.Add(_complectUnitType);
//            checkAmountTypes.Add(_matbaseType);
//            checkAmountTypes.Add(_sostMaterialType);

//            enabledTypes.Add(_CEType);
//            enabledTypes.Add(_partType);
//            enabledTypes.Add(_MOType);
//            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)));
//            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00185-306c-11d8-b4e9-00304f19f545" /*Техпроцесс базовый*/)));
//            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad001ff-306c-11d8-b4e9-00304f19f545" /*Цехозаход*/));
//            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/));
//            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad0017d-306c-11d8-b4e9-00304f19f545" /*Переход*/));
//            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("b3ec04e4-9d56-4494-a57b-766d10cdfe27" /*Группа материалов*/));
//            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00170-306c-11d8-b4e9-00304f19f545" /*Материал базовый*/)));
//            enabledTypes.Add(_asmUnitType);
//            enabledTypes.Add(_complectUnitType);
//            enabledTypes.Add(_zagotType);
//            enabledTypes.Add(_sostMaterialType);

//            string ecoObjCaption = String.Empty;
//            IDBObject ecoObj = session.GetObject(session.EditingContextID, false);
//            IDBObject asmObj = session.GetObject(_asm, false);
//            if (ecoObj != null)
//            {

//                var attr = ecoObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
//                if (attr != null)
//                    document.DocumentName = attr.AsString;

//                attr = asmObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
//                if (attr != null)
//                    document.DocumentName += "__" + attr.AsString;

//                ecoObjCaption = ecoObj.Caption;
//            }
//            else
//            {
//                MessageBox.Show("Не выбран контекст редактирования.", "Ошибка формирования отчета сравнения составов (было-стало)", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                return false;
//            }

//            document.Designation += " " + DateTime.Now.ToString("dd.MM.yyyy, HH:mm:ss");

//            //сборка по извещению
//            IDBObject headerObj = null;
//            //базовая версия
//            IDBObject headerObjBase = null;
//            var list = session.GetObjectIDVersions(objectIDs[0], false);
//            foreach (var item in list)
//            {
//                IDBObject obj = session.GetObject(item, true);
//                if (obj.ModificationID == session.EditingContextID)
//                {
//                    headerObj = obj;
//                }

//                if (obj.IsBaseVersion)
//                    headerObjBase = obj;
//            }

//            string headerObjCaption = String.Empty;
//            if (headerObj == null)
//            {
//                //Если нет созданной по извещению берем ту что подали на вход
//                headerObj = session.GetObject(objectIDs[0], true);
//                headerObjBase = headerObj;

//                headerObjCaption = headerObj.Caption;
//            }

//            if (headerObj == null)
//                throw new Exception("Не найдена версия для данного контекста редактирования");
//            if (headerObjBase == null)
//                throw new Exception("Не найдена базовая версия");
//            bool addBmzFields = true;

//            #region Инициализация запроса к БД

//            List<int> rels = new List<int>();

//            rels.Add(MetaDataHelper.GetRelationTypeID("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/));
//            rels.Add(MetaDataHelper.GetRelationTypeID("cad00023-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));

//            rels.Add(MetaDataHelper.GetRelationTypeID("cad0036b-306c-11d8-b4e9-00304f19f545" /*Изменяется по извещению*/));

//            List<ColumnDescriptor> columns = new List<ColumnDescriptor>();

//            int attrId = (Int32)ObligatoryObjectAttributes.CAPTION;
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//            ColumnNameMapping.FieldName, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstituteInGroup) /*Номер заменителя в группе*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstitutesGroupNo) /*Номер группы заменителя*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("34b64b00-d430-48e4-8692-f52a7485a10a" /*Количество основы (БМЗ)*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("b9ab8a1a-e988-4031-9f30-72a95644830a" /*Количество компонента 1 (БМЗ)*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("8de3b657-868a-47de-8a62-3bbaf7c4994f" /*Количество компонента 2 (БМЗ)*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("54b2e64e-1d4b-4a9d-9633-023f6b6bcd4a" /*Компонент составного материала*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("cad00020-306c-11d8-b4e9-00304f19f545" /*Наименование*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("cad001be-306c-11d8-b4e9-00304f19f545" /*Ссылка на объект*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("cad0038c-306c-11d8-b4e9-00304f19f545" /*Материал*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            attrId = MetaDataHelper.GetAttributeTypeID("cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/);
//            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
//            ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            columns.Add(new ColumnDescriptor(
//            MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/),
//            AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));

//            if (addBmzFields)
//            {
//                attrId = MetaDataHelper.GetAttributeTypeID(new Guid("8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/) /*Признак изготовления*/);
//                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//                ColumnNameMapping.Guid, SortOrders.NONE, 0));
//            }
//            if (addBmzFields)
//            {
//                attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
//                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//                ColumnNameMapping.Guid, SortOrders.NONE, 0));
//            }

//            HybridDictionary tags = new HybridDictionary();
//            DBRecordSetParams dbrsp = new DBRecordSetParams(null,
//            columns != null
//            ? columns.ToArray()
//            : null,
//            0, null,
//            QueryConsts.All);
//            dbrsp.Tags = tags;

//            List<ObjInfoItem> items = new List<ObjInfoItem>();
//            items.Add(new ObjInfoItem(headerObj.ObjectID));

//            #endregion Инициализация запроса к БД

//            IDBEditingContextsService idbECS = session.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;
//            var contextObject = idbECS.GetEditingContextsObject(session.SessionGUID, session.EditingContextID, true, false);
//            if (contextObject != null)
//            {
//                contextObjects = contextObject.GetVersionsID(session.EditingContextID, -1);
//                foreach (var objectid in contextObjects)
//                {
//                    var info = session.GetObjectInfo(objectid);
//                    contextObjectsID.Add(info.ID);
//                }
//            }

//            #region Первый состав по извещению

//            DataTable dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
//            Intermech.SystemGUIDs.filtrationBaseVersions, null, enabledTypes);

//            // Храним пару ид версии объекта + ид. физической величины
//            // те объекты у которых посчитали количество
//            // состав по извещен
//            Dictionary<Tuple<long, long, string>, Item> ecoComposition = GetComposition(headerObj, dt, session);

//            //AddToLog("Первый состав с извещением " + headerObj.ObjectID.ToString());

//            #endregion Первый состав по извещению

//            //сохраняем контекст редактирования
//            long sessionid = session.EditingContextID;
//            session.EditingContextID = 0;

//            #region Второй состав без извещения

//            dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
//            Intermech.SystemGUIDs.filtrationBaseVersions, null, enabledTypes);

//            //AddToLog("Второй состав без извещения " + headerObjBase.ObjectID.ToString());
//            session.EditingContextID = 0;

//            Dictionary<Tuple<long, long, string>, Item> baseComposition = GetComposition(headerObjBase, dt, session);

//            #endregion Второй состав без извещения

//            //возвращаем контекст
//            session.EditingContextID = sessionid;

//            #region Итоговый состав материалов

//            List<ReportRow> resultComposition = new List<ReportRow>();
//            foreach (var resItem in baseComposition)
//            {

//                Item ecoItem;
//                //Находим базовый объект в списке объектов из извещения
//                if (ecoComposition.TryGetValue(resItem.Key, out ecoItem))
//                {
//                    ecoComposition.Remove(resItem.Key);
//                }
//                else
//                {
//                    var emptyItem = new Tuple<long, long, string>(0, resItem.Key.Item2, string.Empty);
//                    if (ecoComposition.TryGetValue(emptyItem, out ecoItem))
//                    {
//                        ecoComposition.Remove(emptyItem);
//                    }
//                }

//                ReportRow material = new ReportRow(resItem.Value, ecoItem);
//                resultComposition.Add(material);
//            }
//            foreach (var ecoItem in ecoComposition.Values)
//            {
//                ReportRow material = new ReportRow(null, ecoItem);
//                resultComposition.Add(material);
//            }

//            resultComposition = resultComposition
//                .Where(e => (e.Amount_base.Value != e.Amount_eco.Value) || e.ReplacementGroupIsChanged || e.ReplacementStatusIsChanged)
//                .ToList();

//            #region Добавление в выборку деталей из заготовок

//            List<ReportRow> zagMaterials = new List<ReportRow>();
//            //материал и список комплектующих(деталей) связанных с ним
//            Dictionary<ReportRow, List<ReportRow>> matPartPairs = new Dictionary<ReportRow, List<ReportRow>>();

//            zagMaterials = resultComposition.Where(e => e.Type == _zagotType)
//                .ToList();


//            foreach (var zgMat in zagMaterials)
//            {
//                var zagBaseComp = zgMat.BaseItem != null ? new Dictionary<long, Item>(zgMat.BaseItem.AssociatedItemsAndSelf) : new Dictionary<long, Item>();
//                var zagEcoComp = zgMat.ECOItem != null ? new Dictionary<long, Item>(zgMat.ECOItem.AssociatedItemsAndSelf) : new Dictionary<long, Item>();

//                List<ReportRow> zags = new List<ReportRow>();

//                foreach (var resItem in zagBaseComp)
//                {
//                    Item ecoItem;
//                    //Находим базовый объект в списке объектов из извещения
//                    if (zagEcoComp.TryGetValue(resItem.Key, out ecoItem))
//                    {
//                        zagEcoComp.Remove(resItem.Key);
//                    }

//                    ReportRow material = new ReportRow(resItem.Value, ecoItem);
//                    zags.Add(material);
//                }
//                foreach (var ecoItem in zagEcoComp.Values)
//                {
//                    ReportRow material = new ReportRow(null, ecoItem);
//                    zags.Add(material);
//                }

//                zags = zags.Where(e => e.Amount_base.Value != e.Amount_eco.Value)
//                   .ToList();

//                //получаем детали от измененных заготовок:
//                List<ReportRow> parts = new List<ReportRow>();

//                foreach (var zag in zags)
//                {
//                    var zagEco = zag.ECOItem;
//                    var zagBase = zag.BaseItem;

//                    var partEco = zagEco != null ? zagEco.RelationsWithParent.First().Parent.RelationsWithParent.First().Parent : null;
//                    var partBase = zagBase != null ? zagBase.RelationsWithParent.First().Parent.RelationsWithParent.First().Parent : null;

//                    //далее пересчитываем общее количество деталей:

//                    Tuple<string, string> exceptionInfo = null;
//                    bool hasContextObjects = false;

//                    if (partEco != null)
//                    {
//                        partEco.SourseOrg = zgMat.SourseOrg;
//                        partEco.WriteToReportForcibly = true;
//                        partEco.LinkToObjId = partEco.ObjectId;

//                        partEco.AmountSum = partEco.GetAmount(false, ref hasContextObjects, ref partEco.HasEmptyAmount, exceptionInfo, null).Values.FirstOrDefault();

//                    }
//                    if (partBase != null)
//                    {
//                        partBase.SourseOrg = zgMat.SourseOrg;
//                        partBase.WriteToReportForcibly = true;
//                        partBase.LinkToObjId = partBase.ObjectId;

//                        partBase.AmountSum = partBase.GetAmount(false, ref hasContextObjects, ref partBase.HasEmptyAmount, exceptionInfo, null).Values.FirstOrDefault();
//                    }
//                    var newPart = new ReportRow(partBase, partEco);
//                    parts.Add(newPart);
//                }

//                if (!matPartPairs.ContainsKey(zgMat))
//                    matPartPairs[zgMat] = parts;
//            }

//            #endregion

//            #endregion Итоговый состав материалов

//            #region Изменение групп заменителей

//            foreach (var item in resultComposition)
//            {
//                if (item.ReplacementGroupIsChanged || item.ReplacementStatusIsChanged)
//                {
//                    //с актуальной на допустимую
//                    if (item.BaseItem.isActualReplacement && item.ECOItem.isPossableReplacement)
//                    {
//                        item.SubstractAmount(amountIsIncrease: false);
//                        var actual = resultComposition.FirstOrDefault(e => e.ECOItem.isActualReplacement && e.ECOItem.ReplacementGroup == item.ECOItem.ReplacementGroup);
//                        item.MaterialCaption += actual != null ? string.Format("\nприменяется взамен {0}", actual.ECOItem.Caption) : "";
//                    }

//                    //с допустимой на актуальную
//                    if (item.BaseItem.isPossableReplacement && item.ECOItem.isActualReplacement)
//                    {
//                        item.SubstractAmount(amountIsIncrease: true);
//                        var possable = resultComposition.FirstOrDefault(e => e.ECOItem.isPossableReplacement && e.ECOItem.ReplacementGroup == item.ECOItem.ReplacementGroup);
//                        item.MaterialCaption += possable != null ? string.Format("\nдопускается замена на {0}", possable.ECOItem.Caption) : "";
//                    }

//                    //с основной на допустимую
//                    if ((!item.BaseItem.isActualReplacement && !item.BaseItem.isPossableReplacement) && item.ECOItem.isPossableReplacement)
//                    {
//                        item.SubstractAmount(amountIsIncrease: false);
//                        var actual = resultComposition.FirstOrDefault(e => e.ECOItem.isActualReplacement && e.ECOItem.ReplacementGroup == item.ECOItem.ReplacementGroup);
//                        item.MaterialCaption += actual != null ? string.Format("\nприменяется взамен {0}", actual.ECOItem.Caption) : "";
//                    }

//                    //с допустимой на основную
//                    if (item.BaseItem.isPossableReplacement && (!item.ECOItem.isActualReplacement && !item.ECOItem.isPossableReplacement))
//                    {
//                        item.SubstractAmount(amountIsIncrease: true);
//                    }

//                    //с основной на актуальную
//                    if ((!item.BaseItem.isActualReplacement && !item.BaseItem.isPossableReplacement) && item.ECOItem.isActualReplacement)
//                    {
//                        var possable = resultComposition.FirstOrDefault(e => e.ECOItem.isPossableReplacement && e.ECOItem.ReplacementGroup == item.ECOItem.ReplacementGroup);
//                        item.MaterialCaption += possable != null ? string.Format("\nдопускается замена на {0}", possable.ECOItem.Caption) : "";
//                    }

//                    //с актуальной на основную
//                    if (item.BaseItem.isActualReplacement && (!item.ECOItem.isActualReplacement && !item.ECOItem.isPossableReplacement))
//                    {
//                        continue;
//                    }
//                }
//            }

//            #endregion

//            List<long> resultCompIds = resultComposition
//                .Select(e => e.LinkToObjId)
//                .Where(e => e > 0)
//                .Distinct()
//                .ToList();

//            resultCompIds.AddRange(
//                resultComposition
//                .Where(e => e.Type == _zagotType)
//                .Select(e => e.MaterialId));

//            resultCompIds = resultCompIds.Distinct().ToList();


//            #region Запись значений атрибутов материалов из соответствующих объектов конструкторского состава

//            List<ReportRow> resultComposition_tech = new List<ReportRow>();
//            resultComposition_tech.AddRange(resultComposition.Where(e => e.IsPurchased));

//            if (resultCompIds.Count > 0)
//            {
//                IDBObjectCollection col = session.GetObjectCollection(-1);

//                List<ConditionStructure> conds = new List<ConditionStructure>();
//                conds.Add(new ConditionStructure((Int32)ObligatoryObjectAttributes.F_OBJECT_ID, RelationalOperators.In,
//                resultCompIds.ToArray(), LogicalOperators.NONE, 0, false));

//                columns = new List<ColumnDescriptor>();

//                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_ID, AttributeSourceTypes.Auto,
//                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

//                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_OBJECT_ID,
//                AttributeSourceTypes.Object,
//                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

//                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.CAPTION,
//                AttributeSourceTypes.Object,
//                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

//                columns.Add(new ColumnDescriptor(
//                MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/),
//                AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));

//                if (addBmzFields)
//                {
//                    attrId = MetaDataHelper.GetAttributeTypeID(new Guid("8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/));
//                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//                    ColumnNameMapping.Guid, SortOrders.NONE, 0));

//                    attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
//                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
//                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
//                }

//                tags = new HybridDictionary();
//                dbrsp = new DBRecordSetParams(conds.ToArray(),
//                columns != null
//                ? columns.ToArray()
//                : null,
//                0, null,
//                QueryConsts.All);
//                dt = col.Select(dbrsp);

//                foreach (DataRow row in dt.Rows)
//                {
//                    long id = Convert.ToInt64(row["F_OBJECT_ID"]);
//                    string name = Convert.ToString(row["CAPTION"]);
//                    string code = Convert.ToString(row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/]);

//                    bool isPurchased = false, isCoop = false;
//                    object attrValue = row["8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/];
//                    if (attrValue is long)
//                    {
//                        isPurchased = Convert.ToInt32(attrValue) != 1 ? true : false;
//                        isCoop = Convert.ToInt32(attrValue) == 3 ? true : false;
//                    }

//                    string sourseOrg = string.Empty;
//                    attrValue = row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/];
//                    if (attrValue is string)
//                        sourseOrg = attrValue.ToString();

//                    var materials = resultComposition.Where(x => (x.LinkToObjId == id) || x.MaterialId == id);
//                    foreach (var mat in materials)
//                    {
//                        if (mat != null && (isPurchased || isCoop))
//                        {
//                            mat.IsPurchased = isPurchased;
//                            mat.MaterialCode = code;
//                            if (mat.Type == _zagotType)
//                                mat.MaterialCaption = name;

//                            resultComposition_tech.Add(mat);
//                            //AddToLog("mat3 " + mat.ToString());
//                        }
//                    }
//                }
//            }

//            #endregion Запись значений атрибутов материалов из соответствующих объектов конструкторского состава


//            List<ReportRow> reportComp = new List<ReportRow>();

//            foreach (var item in resultComposition_tech)
//            {
//                ReportRow repeatItem = null;
//                //когда есть повторение позиции, объединяем с уже записанной
//                foreach (var repItem in reportComp)
//                {
//                    if (repItem.MaterialCaption == item.MaterialCaption && repItem.Type == item.Type && repItem.MeasureUnit == item.MeasureUnit)
//                        repeatItem = repItem;
//                }

//                if (repeatItem != null)
//                    repeatItem.Combine(item);
//                else
//                    reportComp.Add(item);
//            }

//            reportComp.RemoveAll(e => reportComp.Count(i => e.MaterialCaption == i.MaterialCaption && e.MeasureUnit == i.MeasureUnit) > 1);

//            reportComp = reportComp

//                .OrderBy(e => e.MaterialCaption)
//                //.Where(e => (e.Amount1 != e.Amount2))
//                .ToList();

//            reportComp.RemoveAll(e => e.Type == _zagotType);

//            //добавляем новые детали в итоговую выборку
//            foreach (var item in matPartPairs)
//            {
//                reportComp.AddRange(item.Value);
//                reportComp.Add(item.Key);
//            }

//            foreach (var item in reportComp)
//            {
//                //для объектов из других организаций
//                if ((item.SourseOrg != originOrg && (item.Type == _complectUnitType
//                    || MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)).Contains(item.Type)))
//                    || (item.SourseOrg != originOrg && item.Type == _zagotType))
//                    item.MaterialCode += "\nот " + item.SourseOrg;

//                //item.AmountInAsm_eco = item.AmountInAsm_eco.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
//                //item.AmountInAsm_base = item.AmountInAsm_base.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
//            }

//            if (writeTestingData)
//            {
//                foreach (var item in reportComp)
//                {
//                    System.IO.File.AppendAllText(testFile, string.Format("material {0}/{1}/{2}", item.MaterialCaption, item.Amount_base, item.Amount_eco));
//                    foreach (var asm in item.AmountInAsm_base)
//                    {
//                        System.IO.File.AppendAllText(testFile, string.Format(" asm1 {0}/{1}/{2}", asm.Key, asm.Value.Item1.ToString(), asm.Value.Item2.ToString()));
//                    }
//                    foreach (var asm in item.AmountInAsm_base)
//                    {
//                        System.IO.File.AppendAllText(testFile, string.Format(" asm2 {0}/{1}/{2}", asm.Key, asm.Value.Item1.ToString(), asm.Value.Item2.ToString()));
//                    }
//                    System.IO.File.AppendAllText(testFile, "\n");
//                }
//            }

//            #region Формирование отчета

//            // Заполнение шапки
//            //AddToLog("fill header");


//            ReportWordDocument reportWriter = new ReportWordDocument(document, reportComp, ecoObjCaption, headerObjCaption, compliteReport);

//            reportWriter.WriteReport();

//            //AddToLog("Завершили создание отчета ");

//            #endregion Формирование отчета

//            if (_adminError.Length > 0)
//            {
//                IRouterService router = session.GetCustomService(typeof(IRouterService)) as IRouterService;
//                router.CreateMessage(session.SessionGUID, _messageRecepients.ToArray(), "Ошибка формирования отчета сравнения составов (было-стало)", _adminError, session.UserID);
//            }

//            if (_userError.Length > 0)
//                MessageBox.Show(_userError, "Ошибка формирования отчета сравнения составов (было-стало)", MessageBoxButtons.OK, MessageBoxIcon.Warning);

//            return true;
//        }

//        public void AddToLog(string text)
//        {
//            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
//            text = text + Environment.NewLine;
//            System.IO.File.AppendAllText(file, text);
//            //AddToOutputView(text);
//        }

//        private class ReportWordDocument
//        {
//            private ImDocumentData _document;
//            private string _ecoCaption;
//            private string _asmCaption;
//            private List<ReportRow> _resultComposition;
//            private bool _compliteReport;

//            public ReportWordDocument(ImDocumentData document, List<ReportRow> resultComposition, string ecoCaption, string asmCaption, bool completeReport)
//            {
//                _document = document;
//                _ecoCaption = ecoCaption;
//                _asmCaption = asmCaption;
//                _compliteReport = completeReport;
//                _resultComposition = resultComposition;
//            }

//            public void WriteReport()
//            {
//                DocumentTreeNode headerNode = _document.Template.FindNode("Шапка");
//                if (headerNode != null)
//                {
//                    TextData textData = headerNode.FindNode("Извещение") as TextData;
//                    if (textData != null)
//                    {
//                        textData.AssignText(_ecoCaption, false, false, false);
//                    }

//                    textData = headerNode.FindNode("ДСЕ") as TextData;
//                    if (textData != null)
//                    {
//                        textData.AssignText(_asmCaption, false, false, false);
//                    }
//                }

//                DocumentTreeNode docrow = _document.Template.FindNode("Строка");
//                DocumentTreeNode table = _document.FindNode("Таблица");

//                int N = 0;
//                int totalIndex = 0;
//                int rowIndex = 0;

//                foreach (var item in _resultComposition)
//                {
//                    DocumentTreeNode node = docrow.CloneFromTemplate(true, true);
//                    if (_compliteReport)
//                    {

//                        #region Запись "было"

//                        if (item.AmountInAsm_base.Count > 0 || item.AmountInAsm_eco.Count > 0)
//                        {
//                            N++;
//                            totalIndex++;

//                            rowIndex++;
//                            table.AddChildNode(node, false, false);

//                            Write(node, "Индекс", N.ToString());
//                            Write(node, "Код", item.MaterialCode);
//                            if (item.MaterialCaption.Length != item.Caption.Length)
//                            {
//                                Write(node, "Материал", item.MaterialCaption, new CharFormat("arial", 10, CharStyle.Italic));
//                            }
//                            else Write(node, "Материал", item.MaterialCaption);

//                            if (item.Amount_base.Value != 0 || !item.HasEmptyAmount)
//                            {
//                                Write(node, "Всего", Math.Round(item.Amount_base.Value, 3).ToString() + " " + item.MeasureUnit);

//                                if (N == 1)
//                                {
//                                    DocumentTreeNode row2 = node.FindNode("Строка2");
//                                    if (item.AmountInAsm_base.Count != 0)
//                                    {
//                                        AddPresentationRowToCompleteReport(row2, item.AmountInAsm_eco, item.AmountInAsm_base.First(), totalIndex);
//                                    }

//                                    for (int j = 1; j < item.AmountInAsm_base.Count; j++)
//                                    {
//                                        DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
//                                        DocumentTreeNode col2 = node.FindNode("Столбец2");

//                                        totalIndex++;
//                                        col2.AddChildNode(node2, false, false);

//                                        AddPresentationRowToCompleteReport(node2, item.AmountInAsm_eco, item.AmountInAsm_base.ToArray()[j], totalIndex);
//                                    }
//                                }

//                                if (N > 1)
//                                {
//                                    DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", totalIndex));
//                                    if (item.AmountInAsm_base.Count != 0)
//                                    {
//                                        AddPresentationRowToCompleteReport(row2, item.AmountInAsm_eco, item.AmountInAsm_base.First(), totalIndex);
//                                    }

//                                    for (int j = 1; j < item.AmountInAsm_base.Count; j++)
//                                    {
//                                        DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
//                                        DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", rowIndex));

//                                        totalIndex++;
//                                        col2.AddChildNode(node2, false, false);

//                                        AddPresentationRowToCompleteReport(node2, item.AmountInAsm_eco, item.AmountInAsm_base.ToArray()[j], totalIndex);
//                                    }
//                                }
//                            }
//                        }

//                        #endregion Запись "было"

//                        #region Запись "стало"

//                        if (item.AmountInAsm_base.Count > 0 || item.AmountInAsm_eco.Count > 0)
//                        {
//                            totalIndex++;
//                            rowIndex++;
//                            node = docrow.CloneFromTemplate(true, true);

//                            table.AddChildNode(node, false, false);

//                            if (item.Amount_eco.Value != 0 || !item.HasEmptyAmount)
//                            {
//                                Write(node, "Всего", Math.Round(item.Amount_eco.Value, 3).ToString() + " " + item.MeasureUnit);

//                                DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", totalIndex));
//                                DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", rowIndex));
//                                if (item.AmountInAsm_eco.Count != 0)
//                                {
//                                    AddPresentationRowToCompleteReport(row2, item.AmountInAsm_base, item.AmountInAsm_eco.First(), totalIndex);
//                                }

//                                for (int j = 1; j < item.AmountInAsm_eco.Count; j++)
//                                {
//                                    DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
//                                    col2.AddChildNode(node2, false, false);
//                                    totalIndex++;

//                                    AddPresentationRowToCompleteReport(node2, item.AmountInAsm_base, item.AmountInAsm_eco.ToArray()[j], totalIndex);
//                                }
//                            }
//                        }

//                        #endregion Запись "стало"
//                    }
//                    else
//                    {
//                        N++;
//                        table.AddChildNode(node, false, false);

//                        Write(node, "Индекс", N.ToString());
//                        Write(node, "Код", item.MaterialCode);

//                        if (item.MaterialCaption.Length != item.Caption.Length)
//                            Write(node, "Материал", item.MaterialCaption, new CharFormat("arial", 10, CharStyle.Italic));
//                        else Write(node, "Материал", item.MaterialCaption);

//                        if (item.Amount_base.Value != 0)
//                            Write(node, "Было", Convert.ToString(Math.Round(item.Amount_base.Value, 3)));
//                        else
//                            Write(node, "Было", "-");
//                        if (item.Amount_eco.Value != 0)
//                            Write(node, "Будет", Convert.ToString(Math.Round(item.Amount_eco.Value, 3)));
//                        else
//                            Write(node, "Будет", "-");
//                        var diff = Math.Round(item.Amount_eco.Value - item.Amount_base.Value, 3);
//                        Write(node, "Разница", Convert.ToString(diff)/*Convert.ToString(item.Amount2 - item.Amount1)*/);
//                        Write(node, "ЕдИзм", item.MeasureUnit);

//                        if (item.HasEmptyAmount)
//                        {
//                            foreach (DocumentTreeNode child in node.Nodes)
//                            {
//                                SelectNodeInTable(child, BorderStyles.SolidLine, 1, Color.Red, CharStyle.Bold, Color.Red);
//                            }
//                        }
//                    }

//                }
//            }

//            private void SelectNodeInTable(DocumentTreeNode node, BorderStyles borderStyles, float borderWidth, Color borderColor, CharStyle charStyle, Color textColor)
//            {
//                (node as RectangleElement).AssignLeftBorderLine(
//                new BorderLine(borderColor, borderStyles, borderWidth), false);
//                (node as RectangleElement).AssignRightBorderLine(
//                new BorderLine(borderColor, borderStyles, borderWidth), false);
//                (node as RectangleElement).AssignTopBorderLine(
//                new BorderLine(borderColor, borderStyles, borderWidth), false);
//                (node as RectangleElement).AssignBottomBorderLine(
//                new BorderLine(borderColor, borderStyles, borderWidth), false);
//                if (node is TextData)
//                {
//                    if ((node as TextData).CharFormat != null)
//                    {
//                        CharFormat cf = (node as TextData).CharFormat.Clone();
//                        cf.TextColor = textColor;
//                        cf.CharStyle = charStyle;
//                        (node as TextData).SetCharFormat(cf, false, false);
//                    }
//                }
//            }

//            private void AddPresentationRowToCompleteReport(DocumentTreeNode node, Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> amountInAsms_another, KeyValuePair<string, Tuple<MeasuredValue, MeasuredValue>> amountInCurrentAsm, int totalIndex)
//            {
//                bool changedField = false;

//                if (amountInAsms_another.ContainsKey(amountInCurrentAsm.Key))
//                    changedField = amountInCurrentAsm.Value.Item1.Value != amountInAsms_another[amountInCurrentAsm.Key].Item1.Value || amountInCurrentAsm.Value.Item2.Value != amountInAsms_another[amountInCurrentAsm.Key].Item2.Value ? true : false;
//                else
//                    changedField = true;

//                if (totalIndex == 1)
//                {

//                    var charFormat = changedField ? new CharFormat("arial black", 9, CharStyle.Italic | CharStyle.Bold) :
//                        new CharFormat("arial", 9, CharStyle.Regular);
//                    WriteFirstAfterParent(node, "Куда входит", amountInCurrentAsm.Key, charFormat);
//                    WriteFirstAfterParent(node, "Количество вхождений", amountInCurrentAsm.Value.Item1.Caption);
//                    WriteFirstAfterParent(node, "Количество сборок", amountInCurrentAsm.Value.Item2.Caption);

//                    if (changedField)
//                        SelectNodeInTable(node, BorderStyles.SolidLine, 0.5f, Color.Black, CharStyle.Bold, Color.Black);

//                    if (amountInCurrentAsm.Value.Item1.Value == 0 || amountInCurrentAsm.Value.Item2.Value == 0)
//                        SelectNodeInTable(node, BorderStyles.SolidLine, 1, Color.Red, CharStyle.Bold, Color.Red);
//                }
//                else
//                {
//                    var charFormat = changedField ? new CharFormat("arial black", 9, CharStyle.Bold | CharStyle.Italic) :
//                        new CharFormat("arial", 9, CharStyle.Regular);
//                    WriteFirstAfterParent(node, string.Format("Куда входит #{0}", totalIndex), amountInCurrentAsm.Key, charFormat);
//                    WriteFirstAfterParent(node, string.Format("Количество вхождений #{0}", totalIndex), amountInCurrentAsm.Value.Item1.Caption);
//                    WriteFirstAfterParent(node, string.Format("Количество сборок #{0}", totalIndex), amountInCurrentAsm.Value.Item2.Caption);

//                    if (changedField)
//                        SelectNodeInTable(node, BorderStyles.SolidLine, 0.5f, Color.Black, CharStyle.Bold, Color.Black);

//                    if (amountInCurrentAsm.Value.Item1.Value == 0 || amountInCurrentAsm.Value.Item2.Value == 0)
//                        SelectNodeInTable(node, BorderStyles.SolidLine, 1, Color.Red, CharStyle.Bold, Color.Red);
//                }

//            }

//            private void WriteOnlyDiffrenceEntersIn(ReportRow item)
//            {
//                //оставляем только различающиеся элементы EntersInAsm1 и EntersInAsm2
//                var keys1 = item.AmountInAsm_base.Keys.ToList();
//                foreach (var key in keys1)
//                {
//                    Tuple<MeasuredValue, MeasuredValue> value = null;
//                    if (item.AmountInAsm_base.TryGetValue(key, out value))
//                    {
//                        if (value.Item1 == null || value.Item2 == null || item.AmountInAsm_base[key].Item1 == null || item.AmountInAsm_base[key].Item2 == null)
//                        {
//                            item.AmountInAsm_base.Remove(key);
//                            item.AmountInAsm_base.Remove(key);
//                            continue;
//                        }
//                        if (value.Item1.Value == item.AmountInAsm_base[key].Item1.Value && value.Item2.Value == item.AmountInAsm_base[key].Item2.Value)
//                        {
//                            item.AmountInAsm_base.Remove(key);
//                            item.AmountInAsm_base.Remove(key);
//                        }
//                    }
//                }
//            }

//            private DocumentTreeNode AddNode(DocumentTreeNode childNode, string value)
//            {
//                DocumentTreeNode node = childNode.GetDocTreeRoot().Template.FindNode(value).CloneFromTemplate(true, true);
//                childNode.AddChildNode(node, false, false);
//                return node;
//            }

//            private void Write(DocumentTreeNode parent, string tplid, string text)
//            {
//                TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
//                if (td != null)
//                {
//                    td.AssignText(text, false, false, false);
//                }
//            }
//            private void Write(DocumentTreeNode parent, string tplid, string text, CharFormat charFormat = null)
//            {
//                TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
//                if (td != null)
//                {
//                    td.SetCharFormat(charFormat == null ? new CharFormat("arial", 10, CharStyle.Regular) : charFormat, false, false);
//                    td.AssignText(text, false, false, false);
//                }

//            }
//            private void WriteFirstAfterParent(DocumentTreeNode parent, string tplid, string text, CharFormat charFormat = null)
//            {
//                TextData td = parent.FindNode(tplid) as TextData;
//                if (td != null)
//                {
//                    td.SetCharFormat(charFormat == null ? new CharFormat("arial", 10, CharStyle.Regular) : charFormat, false, false);
//                    td.AssignText(text, false, false, false);

//                }
//            }
//        }

//        private class ReportTable
//        {

//        }

//        private class ReportRow
//        {
//            public ReportRow(Item baseItem, Item ecoItem)
//            {
//                this.BaseItem = baseItem;
//                this.ECOItem = ecoItem;
//            }
//            private Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> GetAmountInAsm(Item itemToInit)
//            {
//                Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> entersInAsmToInit = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();
//                if (itemToInit == null)
//                {
//                    return entersInAsmToInit;
//                }

//                //у материала заготовок берем все заготовки
//                var items = new List<Item>();
//                if (itemToInit.ObjectType == MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545" /*Заготовка*/))
//                    items.AddRange(itemToInit.AssociatedItemsAndSelf.Values);
//                else
//                    items.Add(itemToInit);

//                foreach (var item in items)
//                {
//                    foreach (var asm in item.EntersInAsms)
//                    {
//                        var asmAmount = asm.Value.AmountSum;
//                        bool hasContextObjects = false;
//                        var itemAmount = item.GetAmount(false, ref hasContextObjects, ref item.HasEmptyAmount, null, asm.Value).Values.FirstOrDefault();

//                        if (entersInAsmToInit.ContainsKey(asm.Value.Caption) && itemAmount != null)
//                            entersInAsmToInit[asm.Value.Caption].Item1.Add(itemAmount.Clone() as MeasuredValue);
//                        else
//                            entersInAsmToInit[asm.Value.Caption] = new Tuple<MeasuredValue, MeasuredValue>(itemAmount == null ? _emptyAmount : itemAmount, asmAmount == null ? _emptyAmount : asmAmount);
//                    }
//                }

//                return entersInAsmToInit;
//            }

//            private readonly MeasuredValue _emptyAmount = MeasureHelper.ConvertToMeasuredValue("0 шт");

//            public Item BaseItem { get; private set; }
//            public Item ECOItem { get; private set; }

//            public bool IsPurchased
//            {
//                get
//                {
//                    return (ECOItem == null ? BaseItem : ECOItem).isPurchased;
//                }
//                set { }
//            }
//            private long _materialId = 0;
//            public long MaterialId
//            {
//                get
//                {
//                    _materialId = (ECOItem == null ? BaseItem : ECOItem).MaterialId;
//                    return _materialId;
//                }
//            }
//            private string _materialCaption;
//            public string MaterialCaption
//            {
//                get
//                {
//                    if (_materialCaption == null)
//                    {
//                        _materialCaption = (ECOItem == null ? BaseItem : ECOItem).Caption;
//                    }
//                    return _materialCaption;
//                }
//                set
//                {
//                    _materialCaption = value;
//                }
//            }

//            private string _sourseOrg;
//            public string SourseOrg
//            {
//                get
//                {
//                    if (_sourseOrg == null)
//                    {
//                        _sourseOrg = (ECOItem == null ? BaseItem : ECOItem).SourseOrg;
//                    }
//                    return _sourseOrg;
//                }
//                set
//                {
//                    _sourseOrg = value;
//                }
//            }
//            private int _type;
//            public int Type
//            {
//                get
//                {
//                    _type = (ECOItem == null ? BaseItem : ECOItem).ObjectType;
//                    return _type;
//                }
//            }

//            private bool _replacementGroupIsChanged = false;
//            public bool ReplacementGroupIsChanged
//            {
//                get
//                {
//                    if (BaseItem != null && ECOItem != null)
//                        _replacementGroupIsChanged = BaseItem.ReplacementGroup != ECOItem.ReplacementGroup && BaseItem.ReplacementGroup > 0;
//                    return _replacementGroupIsChanged;
//                }
//            }
//            private bool _replacementStatusIsChanged = false;
//            public bool ReplacementStatusIsChanged
//            {
//                get
//                {
//                    if (BaseItem != null && ECOItem != null)
//                        _replacementStatusIsChanged = BaseItem.isActualReplacement != ECOItem.isActualReplacement || BaseItem.isPossableReplacement != ECOItem.isPossableReplacement;
//                    return _replacementStatusIsChanged;
//                }
//            }

//            //public int ReplacementGroup1 = -1;
//            //public bool isActualReplacement1 = false;
//            //public bool isPossableReplacement1 = false;
//            //public int ECOItem.ReplacementGroup = -1;
//            //public bool isActualReplacement2 = false;
//            //public bool isPossableReplacement2 = false;
//            private string _measureUnit = string.Empty;
//            public string MeasureUnit
//            {
//                get
//                {
//                    if (_measureUnit != string.Empty && (ECOItem.AmountSum != null || BaseItem.AmountSum != null))
//                    {
//                        var amountSum = ECOItem.AmountSum == null ? BaseItem.AmountSum : ECOItem.AmountSum;
//                        var descr = MeasureHelper.FindDescriptor(amountSum.MeasureID);
//                        if (descr != null)
//                        {
//                            _measureUnit = descr.ShortName;
//                        }
//                    }
//                    return _measureUnit;
//                }
//            }

//            private string _materialCode = string.Empty;

//            /// <summary>
//            /// Код АМТО
//            /// </summary>
//            public string MaterialCode
//            {
//                get
//                {
//                    if (_materialCode == string.Empty)
//                    {
//                        _materialCode = (ECOItem == null ? BaseItem : ECOItem).MaterialCode;
//                    }
//                    return _materialCode;
//                }
//                set
//                {
//                    if (value == null)
//                        _materialCode = string.Empty;
//                    else
//                        _materialCode = value;
//                }
//            }

//            private long _linkToObjId = 0;
//            public long LinkToObjId
//            {
//                get
//                {
//                    _linkToObjId = (ECOItem == null ? BaseItem : ECOItem).LinkToObjId;
//                    return _linkToObjId;
//                }
//            }

//            private MeasuredValue _amount_base;
//            private MeasuredValue _amount_eco;

//            /// <summary>
//            /// Количество из базовой версии
//            /// </summary>
//            public MeasuredValue Amount_base
//            {
//                get
//                {
//                    if (_amount_base == null)
//                    {
//                        _amount_base = BaseItem == null ? _emptyAmount : BaseItem.AmountSum;
//                    }
//                    return _amount_base;
//                }
//            }

//            /// <summary>
//            /// Количество из версии по извещению
//            /// </summary>
//            public MeasuredValue Amount_eco
//            {
//                get
//                {
//                    if (_amount_eco == null)
//                    {
//                        _amount_eco = ECOItem == null ? _emptyAmount : ECOItem.AmountSum;
//                    }
//                    return _amount_eco;
//                }
//            }


//            private Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> _amountInAsm_base = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();
//            private Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> _amountInAsm_eco = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();

//            /// <summary>
//            /// Вхождения СЕ (название, кол-во, кол-во се)
//            /// </summary>
//            public Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> AmountInAsm_base
//            {
//                get
//                {
//                    if (_amountInAsm_base.Count == 0)
//                    {
//                        _amountInAsm_base = GetAmountInAsm(BaseItem).OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
//                    }
//                    return _amountInAsm_base;
//                }
//                private set
//                {
//                    _amountInAsm_base = value;
//                }
//            }
//            /// <summary>
//            /// Вхождения СЕ (название, кол-во, кол-во се)
//            /// </summary>
//            public Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> AmountInAsm_eco
//            {
//                get
//                {
//                    if (_amountInAsm_eco.Count == 0)
//                    {
//                        _amountInAsm_eco = GetAmountInAsm(ECOItem).OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
//                    }
//                    return _amountInAsm_eco;
//                }
//                private set
//                {
//                    _amountInAsm_eco = value;
//                }
//            }

//            public bool HasEmptyAmount
//            {
//                get
//                {
//                    return (ECOItem == null ? BaseItem : ECOItem).HasEmptyAmount;
//                }
//            }

//            public string Caption
//            {
//                get
//                {
//                    return (ECOItem == null ? BaseItem : ECOItem).Caption;
//                }
//            }

//            public void SubstractAmount(bool amountIsIncrease)
//            {
//                if (amountIsIncrease)
//                {
//                    this.Amount_base.Value = 0;

//                    foreach (var inAsm1 in this.AmountInAsm_base)
//                    {
//                        Tuple<MeasuredValue, MeasuredValue> value = null;
//                        if (this.AmountInAsm_eco.TryGetValue(inAsm1.Key, out value))
//                            inAsm1.Value.Item1.Substract(value.Item1);
//                    }
//                }
//                else
//                {
//                    this.Amount_eco.Value = 0;

//                    foreach (var inAsm2 in this.AmountInAsm_eco)
//                    {
//                        Tuple<MeasuredValue, MeasuredValue> value = null;
//                        if (this.AmountInAsm_base.TryGetValue(inAsm2.Key, out value))
//                            inAsm2.Value.Item1.Substract(value.Item1);
//                    }
//                }
//            }

//            public ReportRow Combine(ReportRow material)
//            {
//                if (material.SourseOrg != this.SourseOrg)
//                    return this;

//                this.Amount_base.Add(material.Amount_base.Clone() as MeasuredValue);
//                this.Amount_eco.Add(material.Amount_eco.Clone() as MeasuredValue);

//                foreach (var eia1 in material.AmountInAsm_base)
//                {
//                    if (!this.AmountInAsm_base.ContainsKey(eia1.Key))
//                        this.AmountInAsm_base.Add(eia1.Key, eia1.Value);
//                    else
//                        this.AmountInAsm_base[eia1.Key].Item1.Add(eia1.Value.Item1);
//                }

//                foreach (var eia2 in material.AmountInAsm_eco)
//                {
//                    if (!this.AmountInAsm_eco.ContainsKey(eia2.Key))
//                        this.AmountInAsm_eco.Add(eia2.Key, eia2.Value);
//                    else
//                        this.AmountInAsm_eco[eia2.Key].Item1.Add(eia2.Value.Item1);
//                }
//                return this;
//            }

//            public override string ToString()
//            {
//                return string.Format("MaterialId={0}; Code={1}; MaterialCaption={2}; Amount_base={3}; Amount_eco={4}",
//                MaterialId, MaterialCode, MaterialCaption, Amount_base, Amount_eco);
//            }
//        }
//    }

//    /// <summary>
//    /// Связь между объектами Item. Через связь записывается исходное значение количества материала
//    /// </summary>
//    public class Relation
//    {
//        public long LinkId;
//        public int RelationTypeId;
//        public Item Child;
//        public Item Parent;
//        public MeasuredValue Amount;
//        public bool HasEmptyAmount = false;

//        public IDictionary<long, MeasuredValue> GetAmount(ref bool hasContextObject, ref bool hasemptyAmountRelations, Tuple<string, string> exceptionInfo, Item endItem)
//        {
//            IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
//            IDictionary<long, MeasuredValue> itemsAmount = Parent.GetAmount(false, ref hasContextObject, ref hasemptyAmountRelations, exceptionInfo, endItem);
//            //Количество инициализируется в методе GetItem
//            // если значение количества пустое, записываем количество у связи, затем возвращаем
//            if (Amount != null)
//            {
//                if (itemsAmount == null || itemsAmount.Count == 0)
//                {
//                    result[MeasureHelper.FindDescriptor(Amount).PhysicalQuantityID] = Amount.Clone() as MeasuredValue;
//                }
//                else
//                {
//                    try
//                    {
//                        foreach (var itemAmount in itemsAmount)
//                        {
//                            itemAmount.Value.Multiply(Amount);
//                            result[MeasureHelper.FindDescriptor(Amount).PhysicalQuantityID] = itemAmount.Value;
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        exceptionInfo = new Tuple<string, string>(this.Child.Caption, ex.Message);
//                        return null;
//                    }
//                }
//            }
//            else
//            {
//                if (HasEmptyAmount)
//                {
//                    hasemptyAmountRelations = true;
//                    if (Parent == null)
//                    {
//                        //AddToLogForLink("Parent == null" + LinkId);
//                    }

//                    if (Child == null)
//                    {
//                        //AddToLogForLink("Child == null" + LinkId);
//                    }

//                    if (Parent != null && Child != null)
//                    {
//                        string text = "Отсутствует количество. Тип связи " +
//                        MetaDataHelper.GetRelationTypeName(RelationTypeId) + " Род. объект " +
//                        Parent.ToString() + " Дочерний объект " + Child.ToString();
//                        string text1 = "Отсутствует количество. " + " Род. объект '" + Parent.Caption +
//                        "' Дочерний объект '" + Child.Caption + "'";

//                        //Script1.AddToOutputView(text1);

//                        #region Перенесем выполнение кода статического метода в текущее место

//                        //IOutputView view = ServicesManager.GetService(typeof(IOutputView)) as IOutputView;
//                        //if (view != null)
//                        //    view.WriteString("Скрипт сравнения составов", text1);
//                        //else
//                        //    AddToLogForLink("view == null");

//                        #endregion Перенесем выполнение кода статического метода в текущее место

//                        //AddToLogForLink(text);
//                    }
//                }

//                if (Child.RelationsWithChild == null || Child.RelationsWithChild.Count == 0)
//                {
//                    //Если это материал и он последний, то его количество равно 0
//                    return result; // null
//                }

//                if (itemsAmount != null)
//                {
//                    result = itemsAmount;
//                }
//            }

//            return result;
//        }

//        public override string ToString()
//        {
//            string res = "Relation=" + LinkId.ToString();
//            if (Parent != null)
//                res = res + " parent= " + Parent.Caption.ToString();
//            if (Child != null)
//                res = res + " child= " + Child.Caption.ToString();
//            res = res + " amount = " + Convert.ToString(Amount);
//            return res;
//        }

//        public void AddToLogForLink(string text)
//        {
//            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
//            text = text + Environment.NewLine;
//            System.IO.File.AppendAllText(file, text);
//            //AddToOutputView(text);
//        }
//    }

//    /// <summary>
//    /// Объект. Хранит ссылки на связи Link; сумму количества материала (получено из связей)
//    /// </summary>
//    public class Item
//    {
//        public override string ToString()
//        {
//            string objType1 = MetaDataHelper.GetObjectTypeName(ObjectType);
//            return string.Format(
//            "Id={3}; Caption={4}; Type={6}; Amount={5}; Code={1}; MaterialId={0}",
//            MaterialId, MaterialCode, Caption, Id, Caption, AmountSum, objType1);
//        }

//        public virtual Item Clone()
//        {
//            Item clone = new Item();
//            clone.MaterialId = MaterialId;
//            clone.MaterialCode = MaterialCode;
//            clone.Id = Id;
//            clone.ObjectId = ObjectId;
//            clone.Caption = Caption;
//            clone.ObjectType = ObjectType;
//            clone.RelationsWithParent = RelationsWithParent;
//            clone.RelationsWithChild = RelationsWithChild;
//            clone._entersInAsms = new Dictionary<long, Item>(_entersInAsms);
//            clone.AmountSum = this.AmountSum != null ? AmountSum.Clone() as MeasuredValue : null;
//            clone.LinkToObjId = LinkToObjId;
//            clone.isPurchased = isPurchased;
//            clone.isPossableReplacement = isPossableReplacement;
//            clone.isActualReplacement = isActualReplacement;
//            clone.ReplacementGroup = ReplacementGroup;
//            clone.isCoop = isCoop;
//            clone.SourseOrg = SourseOrg;
//            clone.WriteToReportForcibly = WriteToReportForcibly;
//            return clone;
//        }
//        public bool WriteToReportForcibly = false;
//        public string SourseOrg;
//        public bool isContextObject = false;
//        public bool isPurchased = false;
//        public bool isCoop = false;
//        public MeasuredValue AmountSum;
//        public long MaterialId;
//        public string MaterialCode;
//        public long Id;
//        public long ObjectId;
//        public string Caption;
//        public long LinkToObjId;
//        public int ObjectType;
//        //группы замен
//        public int ReplacementGroup = -1;
//        public bool isActualReplacement = false;
//        public bool isPossableReplacement = false;

//        private Dictionary<long, Item> _associatedItemsAndSelf = new Dictionary<long, Item>();
//        public Dictionary<long, Item> AssociatedItemsAndSelf
//        {
//            get
//            {
//                if (_associatedItemsAndSelf.Count == 0)
//                    _associatedItemsAndSelf[this.ObjectId] = this.Clone();

//                return _associatedItemsAndSelf;
//            }
//        }
//        /// <summary>
//        /// Связь с первым вхождением в сборку; количество элемента из ближайшей связи
//        /// </summary>
//        public Dictionary<long, Item> EntersInAsms
//        {
//            get
//            {
//                if (_entersInAsms.Keys.Count == 0)
//                {
//                    foreach (var rel in RelationsWithParent)
//                    {
//                        var parent = rel.Parent;
//                        if (parent.ObjectType == MetaDataHelper.GetObjectType(new Guid(SystemGUIDs.objtypeAssemblyUnit)).ObjectTypeID ||
//                            parent.ObjectType == MetaDataHelper.GetObjectType(new Guid("cad00167-306c-11d8-b4e9-00304f19f545" /*Собираемая единица*/)).ObjectTypeID)
//                        {
//                            if (!_entersInAsms.ContainsKey(parent.ObjectId))
//                                _entersInAsms[parent.ObjectId] = parent;
//                        }
//                        else
//                        {
//                            var nextAsms = rel.Parent.EntersInAsms;
//                            foreach (var nextAsm in nextAsms)
//                            {
//                                if (!_entersInAsms.ContainsKey(nextAsm.Key))
//                                    _entersInAsms[nextAsm.Key] = nextAsm.Value;
//                            }
//                        }
//                    }
//                }
//                return _entersInAsms;
//            }
//        }

//        //   public MeasuredValue Amount;
//        /// <summary>
//        /// Связи с дочерними объектами
//        /// </summary>
//        public List<Relation> RelationsWithChild = new List<Relation>();

//        /// <summary>
//        /// Связи с родительскими объектами
//        /// </summary>
//        public List<Relation> RelationsWithParent = new List<Relation>();

//        public bool HasEmptyAmount = false;

//        private Dictionary<long, Item> _entersInAsms = new Dictionary<long, Item>();

//        public IDictionary<long, MeasuredValue> GetAmount(bool checkContextObject, ref bool hasContextObject, ref bool hasemptyAmountRelations, Tuple<string, string> exceptionInfo, Item endItem)
//        {
//            MeasuredValue measuredValue = null;
//            IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
//            if (this.isContextObject) hasContextObject = true;

//            if (exceptionInfo == null)
//                exceptionInfo = new Tuple<string, string>(this.Caption, string.Empty);

//            foreach (Relation relation in RelationsWithParent)
//            {
//                bool hasContextObject1 = true;
//                if (checkContextObject)
//                    hasContextObject1 = false;

//                if (endItem != null)
//                {
//                    Item selectedParentItem = relation.Parent;
//                    //указывает ли элемент на сборку?
//                    if (!selectedParentItem.EntersInAsms.ContainsKey(endItem.ObjectId) && selectedParentItem.ObjectId != endItem.ObjectId)
//                        continue;
//                }

//                IDictionary<long, MeasuredValue> itemsAmount = relation.GetAmount(ref hasContextObject1, ref hasemptyAmountRelations, exceptionInfo, endItem);

//                hasContextObject = hasContextObject | hasContextObject1;
//                if (checkContextObject && !hasContextObject1 && !this.isContextObject) continue;
//                if (itemsAmount == null)
//                {
//                    continue;
//                }

//                foreach (var itemAmount in itemsAmount)
//                {
//                    if (!result.TryGetValue(itemAmount.Key, out measuredValue))
//                    {
//                        result[itemAmount.Key] = itemAmount.Value.Clone() as MeasuredValue;
//                        continue;
//                    }

//                    try
//                    {
//                        measuredValue.Add(itemAmount.Value);
//                    }
//                    catch (Exception ex)
//                    {
//                        exceptionInfo = new Tuple<string, string>(this.Caption, ex.Message);
//                        return null;
//                    }
//                }
//            }

//            return result;
//        }

//        /// <summary>
//        /// Если MaterialId == 0 вернет ObjectId, в противном случае -- MaterialId.
//        /// </summary>
//        /// <returns></returns>
//        public long GetKey()
//        {
//            return this.MaterialId != 0
//            ? this.MaterialId
//            : this.ObjectId;
//        }
//    }

//    /// <summary>
//    /// Составной материал
//    /// </summary>
//    public class ComplexMaterialItem : Item
//    {
//        public MeasuredValue MainComponentAmount;
//        public MeasuredValue Component1Amount;
//        public MeasuredValue Component2Amount;
//    }

//    public class MaterialItem : Item
//    {
//        public ComplexMaterialComponent ComplexMaterialComponentValue = ComplexMaterialComponent.Empty;

//        public enum ComplexMaterialComponent
//        {
//            Component1 = 1,
//            Component2 = 2,
//            ComponentMain = 3,
//            Empty = 4
//        };
//    }
//}